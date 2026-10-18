/**
 * Focus layout for pop-out and meeting windows.
 *
 * The official client's pop-out and meeting windows show only the content of
 * the view. Teams web renders every view (chat, channel, calendar, calls,
 * pre-join) inside one layout with stable regions, so these windows hide the
 * app bar, the list pane and the global header and let the content area fill
 * the window:
 *
 * - `app-layout-area--nav`        the app bar on the left
 * - `app-layout-area--mid-nav`    the list pane (chat list, channel list)
 * - `app-layout-area--title-bar`  the global header (search, back, avatar)
 * - `app-layout-area--main`       the content
 *
 * Hiding is all it takes: the grid tracks of hidden regions collapse, and the
 * header and content take the whole window, in a pop-out (header above the chat)
 * and in a meeting window alike. Positioning the content area by hand was tried
 * and breaks one layout or the other, so none is done. Each rule is scoped with
 * `:has()` on the content area, so if Teams renames the regions the window is
 * left as the full client instead of blank.
 *
 * With a frameless window the header carries the window controls, so it stays.
 */

const { ROLE_MEETING, ROLE_POPOUT } = require("../../popoutWindows/role");

const STYLE_ID = "tfl-popout-layout";

const LAYOUT = '[data-tid="experience-layout"]:has([data-tid="app-layout-area--main"])';

/**
 * @param {{hideTitleBar: boolean}} options
 * @returns {string} the stylesheet for the focus layout
 */
function focusLayoutCss({ hideTitleBar }) {
  const hidden = [
    `${LAYOUT} [data-tid="app-layout-area--nav"]`,
    `${LAYOUT} [data-tid="app-layout-area--mid-nav"]`,
    ...(hideTitleBar ? [`${LAYOUT} [data-tid="app-layout-area--title-bar"]`] : []),
  ];
  return `
${hidden.join(",\n")} {
  display: none !important;
}
`;
}

/**
 * @param {string} role
 * @param {object} config
 * @returns {boolean} whether the focus layout applies to a window
 */
function appliesTo(role, config) {
  return (role === ROLE_POPOUT || role === ROLE_MEETING) && config?.multiWindow?.focusLayout !== false;
}

function install(doc, css) {
  if (doc.getElementById(STYLE_ID)) {
    return;
  }
  const style = doc.createElement("style");
  style.id = STYLE_ID;
  style.textContent = css;
  (doc.head ?? doc.documentElement).appendChild(style);
}

async function init(config, ipcRenderer) {
  const role = await ipcRenderer.invoke("get-window-role");
  if (!appliesTo(role, config)) {
    return;
  }
  const css = focusLayoutCss({ hideTitleBar: config?.frame !== false });
  install(document, css);
  // Teams can replace the head while it boots; put the style back if it goes.
  new MutationObserver(() => install(document, css)).observe(document.documentElement, {
    childList: true,
    subtree: true,
  });
}

module.exports = { init, appliesTo, install, focusLayoutCss, STYLE_ID };
