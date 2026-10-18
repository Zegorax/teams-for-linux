// Pure helpers shared by the main process and the preload. Keep this file free
// of Electron imports: the preload loads it in a renderer.

const ROLE_MAIN = "main";
const ROLE_POPOUT = "popout";
const ROLE_MEETING = "meeting";
const ROLE_ARGUMENT = "--teams-window-role=";

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

/**
 * Reads the window role the main process handed to a preload.
 *
 * @param {string[]} argv - `process.argv` of the renderer
 * @returns {"main"|"popout"|"meeting"}
 */
function roleFromArgv(argv) {
  const entry = (argv || []).find((arg) => arg.startsWith(ROLE_ARGUMENT));
  const role = entry?.slice(ROLE_ARGUMENT.length);
  return role === ROLE_POPOUT || role === ROLE_MEETING ? role : ROLE_MAIN;
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
  MAIN_ONLY_MODULES,
  ROLE_ARGUMENT,
  ROLE_MAIN,
  ROLE_MEETING,
  ROLE_POPOUT,
  moduleRunsInRole,
  roleFromArgv,
};
