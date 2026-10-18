'use strict';

const { describe, it, before, after, beforeEach } = require('node:test');
const assert = require('node:assert');

const electronPath = require.resolve('electron');
const stateKeeperPath = require.resolve('electron-window-state');
const managerPath = require.resolve('../../app/popoutWindows');

const { roleFromArgv, moduleRunsInRole, ROLE_ARGUMENT } = require('../../app/popoutWindows/role');

let created;
let managerModule;

class MockBrowserWindow {
	constructor(options) {
		this.options = options;
		this.id = created.length + 1;
		this.listeners = {};
		this.destroyed = false;
		this.minimized = false;
		this.calls = [];
		this.loaded = [];
		this.webContents = {
			setWindowOpenHandler: (handler) => { this.windowOpenHandler = handler; },
		};
		created.push(this);
	}
	on(event, listener) { this.listeners[event] = listener; }
	isDestroyed() { return this.destroyed; }
	isMinimized() { return this.minimized; }
	restore() { this.calls.push('restore'); }
	show() { this.calls.push('show'); }
	focus() { this.calls.push('focus'); }
	destroy() { this.destroyed = true; this.listeners.closed?.(); }
	loadURL(url, opts) { this.loaded.push([url, opts]); return Promise.resolve(); }
}

before(() => {
	require.cache[electronPath] = {
		id: electronPath, filename: electronPath, loaded: true,
		exports: { BrowserWindow: MockBrowserWindow, nativeTheme: { shouldUseDarkColors: false } },
	};
	require.cache[stateKeeperPath] = {
		id: stateKeeperPath, filename: stateKeeperPath, loaded: true,
		exports: (opts) => ({ x: 10, y: 20, width: opts.defaultWidth, height: opts.defaultHeight, manage() {} }),
	};
	delete require.cache[managerPath];
	managerModule = require(managerPath);
});

after(() => {
	delete require.cache[electronPath];
	delete require.cache[stateKeeperPath];
	delete require.cache[managerPath];
});

beforeEach(() => { created = []; });

describe('window role', () => {
	it('defaults to main and reads popout and meeting from argv', () => {
		assert.equal(roleFromArgv(undefined), 'main');
		assert.equal(roleFromArgv(['electron', 'app']), 'main');
		assert.equal(roleFromArgv([`${ROLE_ARGUMENT}popout`]), 'popout');
		assert.equal(roleFromArgv([`${ROLE_ARGUMENT}meeting`]), 'meeting');
		assert.equal(roleFromArgv([`${ROLE_ARGUMENT}admin`]), 'main', 'unknown roles fall back to main');
	});

	it('keeps app-wide modules in the main window only', () => {
		for (const name of ['trayIconRenderer', 'mqttStatusMonitor', 'meetingStartDetector', 'dockIconRenderer', 'settings']) {
			assert.equal(moduleRunsInRole('main', name), true, name);
			assert.equal(moduleRunsInRole('popout', name), false, name);
			assert.equal(moduleRunsInRole('meeting', name), false, name);
		}
		for (const name of ['speakingIndicator', 'overrideMicConstraints', 'zoom', 'theme', 'popoutChat']) {
			assert.equal(moduleRunsInRole('meeting', name), true, name);
		}
	});
});

describe('URL builders', () => {
	it('builds a chat launcher route for a valid thread on a Teams host', () => {
		assert.equal(
			managerModule.buildChatUrl('https://teams.live.com', '48:notes'),
			'https://teams.live.com/v2/#/l/chat/48%3Anotes/conversations'
		);
		assert.ok(managerModule.buildChatUrl('https://teams.cloud.microsoft/v2/', '19:abc@thread.v2'));
	});

	it('rejects thread ids that could escape the route and non-Teams origins', () => {
		for (const bad of ['', 'a/b', 'a?b=1', 'a#b', 'a b', '../x', undefined, 42, 'x'.repeat(300)]) {
			assert.equal(managerModule.buildChatUrl('https://teams.live.com', bad), null, String(bad));
		}
		assert.equal(managerModule.buildChatUrl('https://evil.example', '48:notes'), null);
		assert.equal(managerModule.buildChatUrl('http://teams.live.com', '48:notes'), null);
		assert.equal(managerModule.buildChatUrl('not a url', '48:notes'), null);
	});

	it('recognises meeting join links but not chat or channel launchers', () => {
		const ok = [
			'https://teams.microsoft.com/l/meetup-join/19%3ameeting_x/0?context=%7B%7D',
			'https://teams.live.com/meet/9876543210?p=abc',
			'https://teams.cloud.microsoft/l/meetup-join/19%3ameeting_x/0',
		];
		const no = [
			'https://teams.microsoft.com/l/chat/0/0?users=a@b.c',
			'https://teams.microsoft.com/l/channel/19%3Ax/General',
			'https://evil.example/l/meetup-join/x',
			'http://teams.microsoft.com/l/meetup-join/x',
			'not a url',
		];
		for (const url of ok) assert.equal(managerModule.isMeetingJoinUrl(url), true, url);
		for (const url of no) assert.equal(managerModule.isMeetingJoinUrl(url), false, url);
	});
});

describe('config helpers', () => {
	it('is on by default and off only when explicitly disabled', () => {
		assert.equal(managerModule.isEnabled({}), true);
		assert.equal(managerModule.isEnabled({ multiWindow: {} }), true);
		assert.equal(managerModule.isEnabled({ multiWindow: { enabled: false } }), false);
	});

	it('opens meetings in a window unless set to main', () => {
		assert.equal(managerModule.meetingTarget({}), 'window');
		assert.equal(managerModule.meetingTarget({ multiWindow: { openMeetingsIn: 'main' } }), 'main');
	});
});

describe('PopoutWindowManager', () => {
	function makeManager(overrides = {}) {
		const opened = [];
		const manager = new managerModule.PopoutWindowManager({
			config: { partition: 'persist:test', chromeUserAgent: 'UA', frame: true, menubar: 'auto', url: 'https://teams.live.com/v2/', ...overrides },
			getMainWindow: () => ({ webContents: { getURL: () => 'https://teams.live.com/v2/' } }),
			handleWindowOpen: (details) => { opened.push(details); return { action: 'deny' }; },
		});
		return { manager, opened };
	}

	it('opens a chat on the shared partition with the pop-out role and no tray ownership', () => {
		const { manager } = makeManager();
		assert.equal(manager.openChat('48:notes'), true);
		const [win] = created;
		assert.equal(win.options.webPreferences.partition, 'persist:test');
		assert.deepEqual(win.options.webPreferences.additionalArguments, [`${ROLE_ARGUMENT}popout`]);
		assert.equal(win.loaded[0][0], 'https://teams.live.com/v2/#/l/chat/48%3Anotes/conversations');
		assert.equal(win.loaded[0][1].userAgent, 'UA');
		assert.equal(manager.size, 1);
	});

	it('focuses the existing window instead of opening a duplicate', () => {
		const { manager } = makeManager();
		manager.openChat('48:notes');
		created[0].minimized = true;
		assert.equal(manager.openChat('48:notes'), true);
		assert.equal(created.length, 1);
		assert.deepEqual(created[0].calls, ['restore', 'show', 'focus']);
		manager.openChat('19:other@thread.v2');
		assert.equal(created.length, 2);
	});

	it('opens meeting links with the meeting role and rejects other URLs', () => {
		const { manager } = makeManager();
		assert.equal(manager.openMeeting('https://teams.live.com/meet/123?p=x'), true);
		assert.deepEqual(created[0].options.webPreferences.additionalArguments, [`${ROLE_ARGUMENT}meeting`]);
		assert.equal(manager.openMeeting('https://teams.live.com/l/chat/0/0?users=a@b.c'), false);
		assert.equal(manager.openMeeting('https://evil.example/meet/1'), false);
		assert.equal(created.length, 1);
	});

	it('does nothing when multiWindow is disabled', () => {
		const { manager } = makeManager({ multiWindow: { enabled: false } });
		assert.equal(manager.openChat('48:notes'), false);
		assert.equal(manager.openMeeting('https://teams.live.com/meet/123'), false);
		assert.equal(created.length, 0);
	});

	it('routes window.open from a pop-out through the shared policy', () => {
		const { manager, opened } = makeManager();
		manager.openChat('48:notes');
		const result = created[0].windowOpenHandler({ url: 'https://example.org/' });
		assert.deepEqual(result, { action: 'deny' });
		assert.equal(opened.length, 1);
	});

	it('forgets a window when it closes and closes every window on request', () => {
		const { manager } = makeManager();
		manager.openChat('48:notes');
		manager.openChat('19:other@thread.v2');
		created[0].listeners.closed();
		assert.equal(manager.size, 1);
		manager.closeAll();
		assert.equal(manager.size, 0);
		assert.equal(created[1].destroyed, true);
	});
});
