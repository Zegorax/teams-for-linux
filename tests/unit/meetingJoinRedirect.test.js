'use strict';

const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert');

const redirect = require('../../app/browser/tools/meetingJoinRedirect');

const KEY = 'tmp.tenant.user.react-web-client.get-user-media-data-bag';
const URL_OK = 'https://teams.live.com/meet/9357262524562?p=abc123';

// Teams stores the record as a JSON string that itself holds JSON.
function record({ url = URL_OK, showPrejoin = true, title = 'Standup', context = 'calendar_grid_event_card_join_button' } = {}) {
	const inner = { id: 'join-1', joinOptions: { context, showPrejoin, title }, coordinates: { meetingData: { meetingUrl: url } } };
	return JSON.stringify(JSON.stringify(inner));
}

function fakeStorage(entries) {
	const keys = Object.keys(entries);
	return { length: keys.length, key: (i) => keys[i], getItem: (k) => entries[k] ?? null };
}

// A document whose pre-join screen, title and buttons can be switched on and off.
function fakeDoc(state) {
	return {
		querySelector: (sel) => {
			if (sel === '[data-tid="calling-prejoin-screen"]') return state.prejoin ? {} : null;
			if (sel === '[data-tid="meeting-header-title"]') return state.title === undefined ? null : { textContent: ` ${state.title} ` };
			if (sel === '[data-tid="prejoin-cancel-button"]') return state.prejoin ? { click: () => { state.cancelled += 1; state.prejoin = false; } } : null;
			if (sel === '[data-tid="prejoin-join-button"]') return state.prejoin ? {} : null;
			return null;
		},
	};
}

describe('meetingJoinRedirect.parseJoinRecord', () => {
	it('reads the link, title and pre-join flag from the double-encoded record', () => {
		assert.deepEqual(redirect.parseJoinRecord(record()), { url: URL_OK, title: 'Standup', showPrejoin: true });
		assert.equal(redirect.parseJoinRecord(record({ showPrejoin: false })).showPrejoin, false);
		assert.equal(redirect.parseJoinRecord(record({ title: '  ' })).title, null);
	});

	it('also reads a record that is encoded once', () => {
		const once = JSON.stringify({ joinOptions: { showPrejoin: true }, coordinates: { meetingData: { meetingUrl: URL_OK } } });
		assert.equal(redirect.parseJoinRecord(once)?.url, URL_OK);
	});

	it('ignores records without a Teams meeting link, and anything that is not JSON', () => {
		assert.equal(redirect.parseJoinRecord(record({ url: 'https://evil.example/meet/1' })), null);
		assert.equal(redirect.parseJoinRecord(record({ url: 'https://teams.live.com/l/chat/0/0?users=a@b.c' })), null);
		assert.equal(redirect.parseJoinRecord(record({ url: 'http://teams.live.com/meet/1' })), null);
		assert.equal(redirect.parseJoinRecord('not json'), null);
		assert.equal(redirect.parseJoinRecord(undefined), null);
		assert.equal(redirect.parseJoinRecord(JSON.stringify({ coordinates: {} })), null);
	});
});

describe('meetingJoinRedirect.readJoinRecord', () => {
	it('finds the record under the Teams key, whatever the user and tenant prefix', () => {
		assert.equal(redirect.readJoinRecord(fakeStorage({ other: record(), [KEY]: record() }))?.url, URL_OK);
		assert.equal(redirect.readJoinRecord(fakeStorage({ 'x.y.react-web-client.get-user-media-data-bag': record() }))?.url, URL_OK);
	});

	it('returns null when there is no usable record', () => {
		assert.equal(redirect.readJoinRecord(fakeStorage({})), null);
		assert.equal(redirect.readJoinRecord(fakeStorage({ other: record() })), null);
		assert.equal(redirect.readJoinRecord(fakeStorage({ [KEY]: 'garbage' })), null);
	});
});

describe('meetingJoinRedirect.redirectPrejoin', () => {
	let invoked;
	const ipc = (result) => ({ invoke: (channel, arg) => { invoked.push([channel, arg]); return Promise.resolve(result); } });
	beforeEach(() => { invoked = []; });

	it('opens the meeting window and cancels the in-place pre-join when the title matches', async () => {
		const state = { prejoin: true, title: 'Standup', cancelled: 0 };
		const ok = await redirect.redirectPrejoin({ doc: fakeDoc(state), storage: fakeStorage({ [KEY]: record() }), ipcRenderer: ipc(true) });
		assert.equal(ok, true);
		assert.deepEqual(invoked, [['popout-open-meeting', URL_OK]]);
		assert.equal(state.cancelled, 1);
	});

	it('leaves the pre-join alone when no window took the link', async () => {
		const state = { prejoin: true, title: 'Standup', cancelled: 0 };
		const ok = await redirect.redirectPrejoin({ doc: fakeDoc(state), storage: fakeStorage({ [KEY]: record() }), ipcRenderer: ipc(false) });
		assert.equal(ok, false);
		assert.equal(invoked.length, 1);
		assert.equal(state.cancelled, 0);
	});

	it('does nothing when the record is for another meeting than the screen shows', async () => {
		const state = { prejoin: true, title: 'Some other meeting', cancelled: 0 };
		const ok = await redirect.redirectPrejoin({ doc: fakeDoc(state), storage: fakeStorage({ [KEY]: record() }), ipcRenderer: ipc(true) });
		assert.equal(ok, false);
		assert.deepEqual(invoked, [], 'a stale record never opens the wrong meeting');
		assert.equal(state.cancelled, 0);
	});

	it('does nothing when there is no record, or the join has no pre-join screen', async () => {
		const state = { prejoin: true, title: 'Standup', cancelled: 0 };
		assert.equal(await redirect.redirectPrejoin({ doc: fakeDoc(state), storage: fakeStorage({}), ipcRenderer: ipc(true) }), false);
		assert.equal(await redirect.redirectPrejoin({ doc: fakeDoc(state), storage: fakeStorage({ [KEY]: record({ showPrejoin: false }) }), ipcRenderer: ipc(true) }), false);
		assert.deepEqual(invoked, []);
	});

	it('does not need the title to match when the record has none', async () => {
		const state = { prejoin: true, title: undefined, cancelled: 0 };
		const ok = await redirect.redirectPrejoin({ doc: fakeDoc(state), storage: fakeStorage({ [KEY]: record({ title: ' ' }) }), ipcRenderer: ipc(true) });
		assert.equal(ok, true);
		assert.equal(state.cancelled, 1);
	});
});

describe('meetingJoinRedirect.init', () => {
	let saved;
	let observers;
	let invoked;

	beforeEach(() => {
		saved = { document: globalThis.document, localStorage: globalThis.localStorage, MutationObserver: globalThis.MutationObserver };
		observers = [];
		invoked = [];
		globalThis.MutationObserver = class {
			constructor(callback) { this.callback = callback; observers.push(this); }
			observe() {}
		};
	});

	afterEach(() => {
		for (const [name, value] of Object.entries(saved)) {
			if (value === undefined) delete globalThis[name]; else globalThis[name] = value;
		}
	});

	const ipc = { invoke: (channel, arg) => { invoked.push([channel, arg]); return Promise.resolve(true); } };
	const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

	it('redirects each time the pre-join screen appears, including a repeat Join with no new record', async () => {
		const state = { prejoin: false, title: 'Standup', cancelled: 0 };
		globalThis.document = { ...fakeDoc(state), documentElement: {} };
		globalThis.localStorage = fakeStorage({ [KEY]: record() });
		redirect.init({}, ipc);
		assert.equal(observers.length, 1);

		for (const round of [1, 2]) {
			state.prejoin = true;
			observers[0].callback();
			await wait(1500);
			assert.equal(invoked.length, round, `join ${round} is redirected`);
			assert.equal(state.cancelled, round);
			assert.equal(state.prejoin, false);
			observers[0].callback();
			await wait(200);
		}
	});

	it('does not react to a pre-join screen that was already there, or to unrelated changes', async () => {
		const state = { prejoin: true, title: 'Standup', cancelled: 0 };
		globalThis.document = { ...fakeDoc(state), documentElement: {} };
		globalThis.localStorage = fakeStorage({ [KEY]: record() });
		redirect.init({}, ipc);
		observers[0].callback();
		await wait(300);
		assert.deepEqual(invoked, []);
	});

	it('does nothing when multiWindow is disabled', () => {
		globalThis.document = { ...fakeDoc({ prejoin: false, cancelled: 0 }), documentElement: {} };
		redirect.init({ multiWindow: { enabled: false } }, ipc);
		assert.equal(observers.length, 0);
	});
});

describe('meetingJoinRedirect timing helpers', () => {
	it('waitForText gives up quietly when the element never has text', async () => {
		const realNow = Date.now;
		let t = 0;
		Date.now = () => (t += 2000);
		try {
			assert.equal(await redirect.waitForText({ querySelector: () => null }, '#x', 3000), null);
		} finally {
			Date.now = realNow;
		}
	});

	it('cancelPrejoin waits for the pre-join to settle, and for the join button too', async () => {
		let clicks = 0;
		let joinPresent = false;
		const doc = {
			querySelector: (sel) => {
				if (sel === '[data-tid="prejoin-cancel-button"]') return { click: () => { clicks += 1; } };
				if (sel === '[data-tid="prejoin-join-button"]') return joinPresent ? {} : null;
				return null;
			},
		};
		const done = redirect.cancelPrejoin(doc);
		await new Promise((resolve) => setTimeout(resolve, 1000));
		assert.equal(clicks, 0, 'no click while the join button is missing');
		joinPresent = true;
		await new Promise((resolve) => setTimeout(resolve, 400));
		assert.equal(clicks, 0, 'no click before the pre-join has settled');
		assert.equal(await done, true);
		assert.equal(clicks, 1);
	});

	it('cancelPrejoin gives up quietly when there is no pre-join screen', async () => {
		const realNow = Date.now;
		let t = 0;
		Date.now = () => (t += 4000);
		try {
			assert.equal(await redirect.cancelPrejoin({ querySelector: () => null }), false);
		} finally {
			Date.now = realNow;
		}
	});
});
