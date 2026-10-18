'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');

const { threadIdFromListTarget, currentThreadId } = require('../../app/browser/tools/popoutChat');

// Minimal stand-in for the few DOM members the module reads.
function el(id, children = [], parent = null) {
	const node = { nodeType: 1, id, children, parentElement: parent };
	for (const child of children) child.parentElement = node;
	node.querySelectorAll = (selector) => {
		const prefix = /\[id\^="([^"]+)"\]/.exec(selector)[1];
		const found = [];
		const walk = (n) => { for (const c of n.children) { if (c.id.startsWith(prefix)) found.push(c); walk(c); } };
		walk(node);
		return found;
	};
	return node;
}

describe('popoutChat thread id lookup', () => {
	it('finds the thread from a click on the title, time or row of a single list item', () => {
		const title = el('title-chat-list-item_48:notes');
		const time = el('time-chat-list-item_48:notes');
		const row = el('', [title, time]);
		const list = el('', [row, el('', [el('title-chat-list-item_19:b@thread.v2')])]);
		assert.equal(threadIdFromListTarget(title), '48:notes');
		assert.equal(threadIdFromListTarget(time), '48:notes');
		assert.equal(threadIdFromListTarget(row), '48:notes');
		assert.ok(list.children.length === 2);
	});

	it('returns null when the click is outside any single item', () => {
		const a = el('title-chat-list-item_a');
		const b = el('title-chat-list-item_b');
		const gap = el('');
		const list = el('', [el('', [a]), el('', [b]), gap]);
		assert.equal(threadIdFromListTarget(list), null);
		assert.equal(threadIdFromListTarget(gap), null);
		assert.equal(threadIdFromListTarget(null), null);
	});

	it('reads the open chat from the chat header id', () => {
		const doc = { querySelector: () => ({ id: 'chat-header-48:notes' }) };
		assert.equal(currentThreadId(doc), '48:notes');
		assert.equal(currentThreadId({ querySelector: () => null }), null);
	});
});
