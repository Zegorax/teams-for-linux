// Pure helpers shared by the main process and the preload. Keep this file free
// of Electron imports: the preload loads it in a renderer.
const { isTeamsHost } = require("../helpers/teamsHosts");

const ROLE_MAIN = "main";
const ROLE_POPOUT = "popout";
const ROLE_MEETING = "meeting";

// Modules that own app-wide state (tray, badge, MQTT, dock). They run in the
// main window only, so a pop-out neither duplicates nor overwrites that state.
const MAIN_ONLY_MODULES = new Set([
  "settings",
  "trayIconRenderer",
  "mqttStatusMonitor",
  "meetingStartDetector",
  "dockIconRenderer",
  "customStickers",
]);

// A renderer asks for a pop-out by calling `window.open` with a frame name of
// the form `tfl-popout:<kind>:<key>`. The window then has an opener on the same
// host, which is what makes Teams treat it as a secondary window. The main
// process recognises the name in `setWindowOpenHandler`.
const POPOUT_FRAME_PREFIX = "tfl-popout:";
const KIND_CHAT = "chat";

// Teams thread ids look like `19:abc...@thread.v2`, `19:<guid>_<guid>@unq.gbl.spaces`
// or `48:notes`. The allowlist keeps anything that could break out of the
// hash route (`/`, `?`, `#`, whitespace) out of the URL we build.
const THREAD_ID_PATTERN = /^[0-9A-Za-z:@._~-]{1,256}$/;

/**
 * Builds the URL that opens a chat thread in a fresh Teams document. The
 * `/l/chat/<thread>/conversations` launcher route is the one the SPA resolves
 * in page (see mainAppWindow/deepLinkRouter.js).
 *
 * @param {string} origin - Teams origin such as `https://teams.live.com`
 * @param {string} threadId
 * @returns {string|null} null when the origin or thread id is not acceptable
 */
function buildChatUrl(origin, threadId) {
  if (typeof threadId !== "string" || !THREAD_ID_PATTERN.test(threadId)) {
    return null;
  }
  let parsed;
  try {
    parsed = new URL(origin);
  } catch {
    return null;
  }
  if (parsed.protocol !== "https:" || !isTeamsHost(parsed.hostname)) {
    return null;
  }
  return `${parsed.origin}/v2/#/l/chat/${encodeURIComponent(threadId)}/conversations`;
}

/**
 * @param {string} kind
 * @param {string} key
 * @returns {string} frame name for `window.open`
 */
function popoutFrameName(kind, key) {
  return `${POPOUT_FRAME_PREFIX}${kind}:${key}`;
}

/**
 * @param {string} frameName
 * @returns {{kind: string, key: string}|null}
 */
function parsePopoutFrameName(frameName) {
  if (typeof frameName !== "string" || !frameName.startsWith(POPOUT_FRAME_PREFIX)) {
    return null;
  }
  const rest = frameName.slice(POPOUT_FRAME_PREFIX.length);
  const split = rest.indexOf(":");
  if (split <= 0 || split === rest.length - 1) {
    return null;
  }
  return { kind: rest.slice(0, split), key: rest.slice(split + 1) };
}

/**
 * @param {string} role
 * @param {string} moduleName
 * @returns {boolean} whether a browser module runs in a window of this role
 */
function moduleRunsInRole(role, moduleName) {
  return role === ROLE_MAIN || !MAIN_ONLY_MODULES.has(moduleName);
}

module.exports = {
  KIND_CHAT,
  MAIN_ONLY_MODULES,
  POPOUT_FRAME_PREFIX,
  ROLE_MAIN,
  ROLE_MEETING,
  ROLE_POPOUT,
  buildChatUrl,
  moduleRunsInRole,
  parsePopoutFrameName,
  popoutFrameName,
};
