'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');

const layout = require('../../app/browser/tools/popoutLayout');

describe('popoutLayout', () => {
	it('applies to pop-out windows only, unless focusLayout is turned off', () => {
		assert.equal(layout.appliesTo('popout', {}), true);
		assert.equal(layout.appliesTo('popout', { multiWindow: { focusLayout: true } }), true);
		assert.equal(layout.appliesTo('popout', { multiWindow: { focusLayout: false } }), false);
		assert.equal(layout.appliesTo('main', {}), false, 'the main window keeps the full layout');
		assert.equal(layout.appliesTo('meeting', {}), false, 'meeting windows are not reduced yet');
	});

	it('hides the app bar and list pane and widens the content area', () => {
		const css = layout.FOCUS_LAYOUT_CSS;
		assert.match(css, /app-layout-area--nav/);
		assert.match(css, /app-layout-area--mid-nav/);
		assert.match(css, /app-layout-area--main[^{]*\{[^}]*width: 100% !important/);
		assert.match(css, /display: none !important/);
	});

	it('only reduces the layout when the content area exists, so a Teams rename leaves the full client', () => {
		const rules = layout.FOCUS_LAYOUT_CSS.split('}').map((r) => r.trim()).filter(Boolean);
		assert.ok(rules.length >= 2);
		for (const rule of rules) {
			assert.match(rule, /\[data-tid="experience-layout"\]:has\(\[data-tid="app-layout-area--main"\]\)/);
		}
	});

	it('installs one style element and does not duplicate it', () => {
		const head = { children: [], appendChild(el) { this.children.push(el); } };
		const doc = {
			head,
			documentElement: head,
			getElementById: (id) => head.children.find((c) => c.id === id) || null,
			createElement: () => ({ id: '', textContent: '' }),
		};
		layout.install(doc);
		layout.install(doc);
		assert.equal(head.children.length, 1);
		assert.equal(head.children[0].id, layout.STYLE_ID);
		assert.equal(head.children[0].textContent, layout.FOCUS_LAYOUT_CSS);
	});
});
