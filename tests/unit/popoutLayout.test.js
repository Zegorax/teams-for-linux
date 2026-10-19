'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');

const layout = require('../../app/browser/tools/popoutLayout');

describe('popoutLayout', () => {
	it('applies to pop-out and meeting windows, unless focusLayout is turned off', () => {
		assert.equal(layout.appliesTo('popout', {}), true);
		assert.equal(layout.appliesTo('meeting', {}), true);
		assert.equal(layout.appliesTo('popout', { multiWindow: { focusLayout: true } }), true);
		assert.equal(layout.appliesTo('popout', { multiWindow: { focusLayout: false } }), false);
		assert.equal(layout.appliesTo('meeting', { multiWindow: { focusLayout: false } }), false);
		assert.equal(layout.appliesTo('main', {}), false, 'the main window keeps the full layout');
	});

	it('hides the app bar, list pane and header and leaves the grid to collapse', () => {
		const css = layout.focusLayoutCss({ hideTitleBar: true });
		for (const region of ['app-layout-area--nav', 'app-layout-area--mid-nav', 'app-layout-area--title-bar']) {
			assert.match(css, new RegExp(region));
		}
		assert.match(css, /display: none !important/);
		assert.doesNotMatch(css, /position:|inset:/, 'no hand positioning: it breaks one layout or the other');
		assert.doesNotMatch(css, /app-layout-area--main"\]\s*[,{]/, 'the content area itself is never targeted');
	});

	it('keeps the header when the window is frameless, since it carries the window controls', () => {
		const css = layout.focusLayoutCss({ hideTitleBar: false });
		assert.doesNotMatch(css, /app-layout-area--title-bar/);
		assert.match(css, /app-layout-area--nav/, 'the app bar and list pane are still hidden');
	});

	it('only reduces the layout when the content area exists, so a Teams rename leaves the full client', () => {
		const rules = layout.focusLayoutCss({ hideTitleBar: true }).split('}').map((r) => r.trim()).filter(Boolean);
		assert.ok(rules.length >= 1);
		for (const rule of rules) {
			for (const selector of rule.split('{')[0].split(',')) {
				assert.match(selector.trim(), /^\[data-tid="experience-layout"\]:has\(\[data-tid="app-layout-area--main"\]\)/);
			}
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
		layout.install(doc, 'a{}');
		layout.install(doc, 'a{}');
		assert.equal(head.children.length, 1);
		assert.equal(head.children[0].id, layout.STYLE_ID);
		assert.equal(head.children[0].textContent, 'a{}');
	});
});
