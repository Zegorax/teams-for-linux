/**
 * Focus layout for pop-out windows.
 *
 * The official client's pop-out shows only the content of the view. Teams web
 * renders every view (chat, channel, calendar, calls) inside one layout with
 * stable regions, so a pop-out hides the app bar and the list pane and lets the
 * content area fill the window:
 *
 * - `app-layout-area--nav`      the app bar on the left
 * - `app-layout-area--mid-nav`  the list pane (chat list, channel list)
 * - `app-layout-area--main`     the content
 *
 * Each rule is scoped with `:has()` on the content area, so if Teams renames
 * the regions the page is left as the full client instead of going blank.
 */

const { ROLE_POPOUT } = require("../../popoutWindows/role");

const STYLE_ID = "tfl-popout-layout";

const LAYOUT = '[data-tid="experience-layout"]:has([data-tid="app-layout-area--main"])';

const FOCUS_LAYOUT_CSS = `
${LAYOUT} [data-tid="app-layout-area--nav"],
${LAYOUT} [data-tid="app-layout-area--mid-nav"] {
  display: none !important;
}
${LAYOUT} [data-tid="app-layout-area--main"] {
  left: 0 !important;
  width: 100% !important;
  max-width: none !important;
}
`;

/**
 * @param {string} role
 * @param {object} config
 * @returns {boolean} whether the focus layout applies to a window
 */
function appliesTo(role, config) {
  return role === ROLE_POPOUT && config?.multiWindow?.focusLayout !== false;
}

function install(doc) {
  if (doc.getElementById(STYLE_ID)) {
    return;
  }
  const style = doc.createElement("style");
  style.id = STYLE_ID;
  style.textContent = FOCUS_LAYOUT_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

async function init(config, ipcRenderer) {
  const role = await ipcRenderer.invoke("get-window-role");
  if (!appliesTo(role, config)) {
    return;
  }
  install(document);
  // Teams can replace the head while it boots; put the style back if it goes.
  new MutationObserver(() => install(document)).observe(document.documentElement, {
    childList: true,
    subtree: true,
  });
}

module.exports = { init, appliesTo, install, FOCUS_LAYOUT_CSS, STYLE_ID };
