'use strict';

const { describe, it, before, after, beforeEach } = require('node:test');
const assert = require('node:assert');

const electronPath = require.resolve('electron');
const stateKeeperPath = require.resolve('electron-window-state');
const managerPath = require.resolve('../../app/popoutWindows');

const { moduleRunsInRole, popoutFrameName, parsePopoutFrameName } = require('../../app/popoutWindows/role');

let created;
let managerModule;

class MockBrowserWindow {
	constructor(options) {
		this.options = options;
		this.loaded = [];
		this.id = created.length + 1;
		this.listeners = {};
		this.destroyed = false;
		this.minimized = false;
		this.calls = [];
		this.webContents = {
			id: 100 + this.id,
			handlers: {},
			setWindowOpenHandler: (handler) => { this.windowOpenHandler = handler; },
			on: (event, listener) => { this.webContents.handlers[event] = listener; },
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

describe('pop-out frame names', () => {
	it('round-trips kind and key, keeping colons in the key', () => {
		const name = popoutFrameName('chat', '19:abc@thread.v2');
		assert.deepEqual(parsePopoutFrameName(name), { kind: 'chat', key: '19:abc@thread.v2' });
	});

	it('ignores names that are not pop-out requests or have no key', () => {
		for (const name of [undefined, '', '_blank', 'tfl-popout:', 'tfl-popout:chat:', 'tfl-popout::x', 'other:chat:x']) {
			assert.equal(parsePopoutFrameName(name), null, String(name));
		}
	});
});

describe('PopoutWindowManager', () => {
	const ORIGIN = 'https://teams.live.com';
	const chatDetails = (thread) => ({
		url: managerModule.buildChatUrl(ORIGIN, thread),
		frameName: popoutFrameName('chat', thread),
	});

	function makeManager(overrides = {}) {
		const fallback = [];
		const manager = new managerModule.PopoutWindowManager({
			config: { frame: true, menubar: 'auto', onNewWindowOpenMeetupJoinUrlInApp: true, ...overrides },
			fallbackWindowOpen: (details) => { fallback.push(details); return { action: 'deny', fallback: true }; },
		});
		return { manager, fallback };
	}

	function open(manager, details) {
		const response = manager.handleWindowOpen(details);
		if (response?.action !== 'allow') return response;
		const win = new MockBrowserWindow();
		manager.adopt(win, details);
		return { response, win };
	}

	it('allows a chat pop-out with the pop-out role and the opener preserved', () => {
		const { manager } = makeManager();
		const { response, win } = open(manager, chatDetails('48:notes'));
		assert.equal(response.action, 'allow');
		const prefs = response.overrideBrowserWindowOptions.webPreferences;
		assert.equal(prefs.additionalArguments, undefined, 'the role is looked up over IPC, not passed as an argument');
		assert.equal(prefs.contextIsolation, false);
		assert.equal(prefs.partition, undefined, 'the child inherits the opener session');
		assert.equal(win.__teamsRole, 'popout');
		assert.equal(manager.roleOf(win.webContents.id), 'popout');
		assert.equal(manager.roleOf(9999), 'main', 'unknown renderers are the main window');
		assert.equal(manager.size, 1);
	});

	it('focuses and denies a repeat request instead of opening a duplicate', () => {
		const { manager } = makeManager();
		const { win } = open(manager, chatDetails('48:notes'));
		win.minimized = true;
		assert.deepEqual(manager.handleWindowOpen(chatDetails('48:notes')), { action: 'deny' });
		assert.deepEqual(win.calls, ['restore', 'show', 'focus']);
		assert.equal(manager.handleWindowOpen(chatDetails('19:other@thread.v2')).action, 'allow');
	});

	it('rejects a pop-out name whose URL is not the chat launcher route', () => {
		const { manager } = makeManager();
		for (const url of [
			'https://evil.example/v2/#/l/chat/48%3Anotes/conversations',
			'https://teams.live.com/v2/#/l/chat/other/conversations',
			'https://teams.live.com/somewhere/else',
		]) {
			assert.equal(manager.handleWindowOpen({ url, frameName: popoutFrameName('chat', '48:notes') }), null, url);
		}
		assert.equal(manager.handleWindowOpen({ url: chatDetails('x').url, frameName: popoutFrameName('unknown', 'x') }), null);
	});

	it('opens a meeting link in a meeting window of its own, without an opener', () => {
		const { manager } = makeManager({ partition: 'persist:test', chromeUserAgent: 'UA' });
		const url = 'https://teams.live.com/meet/123?p=x';
		assert.deepEqual(manager.handleWindowOpen({ url, frameName: '' }), { action: 'deny' });
		assert.equal(created.length, 1);
		const win = created[0];
		assert.equal(win.options.webPreferences.partition, 'persist:test');
		assert.equal(win.options.webPreferences.contextIsolation, false);
		assert.deepEqual(win.loaded, [[url, { userAgent: 'UA' }]]);
		assert.equal(win.__teamsRole, 'meeting');
		assert.equal(manager.roleOf(win.webContents.id), 'meeting');
		assert.equal(manager.size, 1);
	});

	it('focuses an open meeting window instead of opening another', () => {
		const { manager } = makeManager();
		const url = 'https://teams.live.com/meet/123?p=x';
		manager.handleWindowOpen({ url, frameName: '' });
		created[0].minimized = true;
		assert.deepEqual(manager.handleWindowOpen({ url, frameName: '' }), { action: 'deny' });
		assert.equal(created.length, 1);
		assert.deepEqual(created[0].calls, ['restore', 'show', 'focus']);
	});

	it('opens a meeting link from outside Teams (a deep link) in a meeting window', () => {
		const { manager } = makeManager({ partition: 'persist:test', chromeUserAgent: 'UA' });
		const url = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_x/0?context=%7B%7D';
		assert.equal(manager.openMeeting(url), true);
		assert.equal(created.length, 1);
		assert.deepEqual(created[0].loaded, [[url, { userAgent: 'UA' }]]);
		created[0].minimized = true;
		assert.equal(manager.openMeeting(url), true, 'a repeat focuses the open window');
		assert.equal(created.length, 1);
		assert.deepEqual(created[0].calls, ['restore', 'show', 'focus']);
	});

	it('does not take deep links that are not meeting joins, or when told to use the main window', () => {
		assert.equal(makeManager().manager.openMeeting('https://teams.live.com/l/chat/0/0?users=a@b.c'), false);
		assert.equal(makeManager().manager.openMeeting('https://evil.example/meet/1'), false);
		assert.equal(makeManager({ multiWindow: { openMeetingsIn: 'main' } }).manager.openMeeting('https://teams.live.com/meet/1'), false);
		assert.equal(makeManager({ multiWindow: { enabled: false } }).manager.openMeeting('https://teams.live.com/meet/1'), false);
		assert.equal(created.length, 0);
	});

	it('leaves other links alone', () => {
		const { manager } = makeManager();
		assert.equal(manager.handleWindowOpen({ url: 'https://teams.live.com/l/chat/0/0?users=a@b.c', frameName: '' }), null);
		assert.equal(manager.handleWindowOpen({ url: 'https://evil.example/meet/1', frameName: '' }), null);
		assert.equal(created.length, 0);
	});

	it('keeps meeting links out of a window when told to use the main window or the in-app option is off', () => {
		const url = 'https://teams.live.com/meet/123';
		assert.equal(makeManager({ multiWindow: { openMeetingsIn: 'main' } }).manager.handleWindowOpen({ url, frameName: '' }), null);
		assert.equal(makeManager({ onNewWindowOpenMeetupJoinUrlInApp: false }).manager.handleWindowOpen({ url, frameName: '' }), null);
	});

	it('does nothing when multiWindow is disabled', () => {
		const { manager } = makeManager({ multiWindow: { enabled: false } });
		assert.equal(manager.handleWindowOpen(chatDetails('48:notes')), null);
		assert.equal(manager.handleWindowOpen({ url: 'https://teams.live.com/meet/123', frameName: '' }), null);
	});

	it('does not adopt a window it did not allow', () => {
		const { manager } = makeManager();
		manager.adopt(new MockBrowserWindow({}), chatDetails('48:notes'));
		assert.equal(manager.size, 0);
	});

	it('installs the policy on the main web contents and on adopted windows, falling back for other links', () => {
		const { manager, fallback } = makeManager();
		const mainContents = new MockBrowserWindow().webContents;
		manager.attach(mainContents);
		const { win } = open(manager, chatDetails('48:notes'));
		assert.equal(typeof win.windowOpenHandler, 'function', 'adopted window gets the policy');
		const other = win.windowOpenHandler({ url: 'https://example.org/', frameName: '' });
		assert.deepEqual(other, { action: 'deny', fallback: true });
		assert.equal(fallback.length, 1);
	});

	it('adopts a created window through did-create-window', () => {
		const { manager } = makeManager();
		const mainContents = new MockBrowserWindow().webContents;
		manager.attach(mainContents);
		const details = chatDetails('48:notes');
		manager.handleWindowOpen(details);
		mainContents.handlers['did-create-window'](new MockBrowserWindow(), details);
		assert.equal(manager.size, 1);
	});

	it('forgets a window when it closes and closes every window on request', () => {
		const { manager } = makeManager();
		const first = open(manager, chatDetails('48:notes')).win;
		const second = open(manager, chatDetails('19:other@thread.v2')).win;
		first.listeners.closed();
		assert.equal(manager.size, 1);
		manager.closeAll();
		assert.equal(manager.size, 0);
		assert.equal(second.destroyed, true);
	});
});
