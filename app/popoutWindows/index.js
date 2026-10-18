const { BrowserWindow, nativeTheme } = require("electron");
const path = require("node:path");
const windowStateKeeper = require("electron-window-state");
const { isTeamsHost } = require("../helpers/teamsHosts");

const { ROLE_ARGUMENT, ROLE_MEETING, ROLE_POPOUT, roleFromArgv } = require("./role");

const MIN_WIDTH = 320;
const MIN_HEIGHT = 240;
const CASCADE_STEP = 30;

// Teams thread ids look like `19:abc...@thread.v2`, `19:<guid>_<guid>@unq.gbl.spaces`
// or `48:notes`. The allowlist keeps anything that could break out of the
// hash route (`/`, `?`, `#`, whitespace) out of the URL we build.
const THREAD_ID_PATTERN = /^[0-9A-Za-z:@._~-]{1,256}$/;

// Join links only. `meetupJoinRegEx` also matches chat and channel launchers,
// which must keep their existing behaviour.
const MEETING_PATH_PATTERN = /^\/(?:l\/meetup-join\/|meet\/|v2\/meet\/)/;

/**
 * Whether the multiWindow feature is on. It is on unless explicitly disabled.
 *
 * @param {object} config
 * @returns {boolean}
 */
function isEnabled(config) {
  return config?.multiWindow?.enabled !== false;
}

/**
 * @param {object} config
 * @returns {"window"|"main"}
 */
function meetingTarget(config) {
  return config?.multiWindow?.openMeetingsIn === "main" ? "main" : "window";
}

/**
 * Whether a URL is a Teams meeting join link on a Teams host.
 *
 * @param {string} url
 * @returns {boolean}
 */
function isMeetingJoinUrl(url) {
  try {
    const parsed = new URL(url);
    return (
      parsed.protocol === "https:" &&
      isTeamsHost(parsed.hostname) &&
      MEETING_PATH_PATTERN.test(parsed.pathname)
    );
  } catch {
    return false;
  }
}

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
 * Owns the secondary Teams windows: chat pop-outs and meeting windows.
 *
 * Every window runs the full Teams web app on the same session partition as
 * the main window, so it is signed in without another prompt. A window is a
 * peer that can be closed on its own; it never hides to the tray and never
 * quits the app. The main window owns the tray, the badge and the app
 * lifetime.
 */
class PopoutWindowManager {
  #config;
  #iconImage;
  #getMainWindow;
  #handleWindowOpen;
  #windows = new Map();
  #cascade = 0;

  /**
   * @param {object} deps
   * @param {object} deps.config - startup configuration
   * @param {() => Electron.BrowserWindow|null} deps.getMainWindow
   * @param {(details: object) => object} deps.handleWindowOpen - window-open
   *   policy shared with the main window (external links, auth popups)
   * @param {Electron.NativeImage} [deps.iconImage]
   */
  constructor({ config, getMainWindow, handleWindowOpen, iconImage }) {
    this.#config = config;
    this.#getMainWindow = getMainWindow;
    this.#handleWindowOpen = handleWindowOpen;
    this.#iconImage = iconImage;
  }

  get size() {
    return this.#windows.size;
  }

  list() {
    return [...this.#windows.entries()].map(([key, win]) => ({
      key,
      id: win.id,
      role: win.__teamsRole,
    }));
  }

  /**
   * Opens a chat thread in its own window, or focuses the window that already
   * shows it.
   *
   * @param {string} threadId
   * @returns {boolean} true when a window was opened or focused
   */
  openChat(threadId) {
    if (!isEnabled(this.#config)) {
      return false;
    }
    const url = buildChatUrl(this.#teamsOrigin(), threadId);
    if (!url) {
      console.warn("[POPOUT] Rejected chat pop-out request");
      return false;
    }
    return this.#openOrFocus(`chat:${threadId}`, ROLE_POPOUT, url);
  }

  /**
   * Opens a meeting join link in its own window.
   *
   * @param {string} url
   * @returns {boolean} true when a window was opened or focused
   */
  openMeeting(url) {
    if (!isEnabled(this.#config) || !isMeetingJoinUrl(url)) {
      return false;
    }
    return this.#openOrFocus(`meeting:${url}`, ROLE_MEETING, url);
  }

  closeAll() {
    for (const win of [...this.#windows.values()]) {
      if (!win.isDestroyed()) {
        win.destroy();
      }
    }
    this.#windows.clear();
  }

  #teamsOrigin() {
    const main = this.#getMainWindow();
    const candidates = [main?.webContents?.getURL?.(), this.#config.url];
    for (const candidate of candidates) {
      try {
        const parsed = new URL(candidate);
        if (parsed.protocol === "https:" && isTeamsHost(parsed.hostname)) {
          return parsed.origin;
        }
      } catch {
        // try the next candidate
      }
    }
    return "https://teams.microsoft.com";
  }

  #openOrFocus(key, role, url) {
    const existing = this.#windows.get(key);
    if (existing && !existing.isDestroyed()) {
      if (existing.isMinimized()) {
        existing.restore();
      }
      existing.show();
      existing.focus();
      return true;
    }
    const win = this.#create(key, role);
    // The SPA taking over the route aborts the first load in normal use, so a
    // rejected promise here is not a failure.
    win.loadURL(url, { userAgent: this.#config.chromeUserAgent }).catch((error) => {
      console.debug("[POPOUT] Navigation interrupted", { code: error?.code ?? error?.errno });
    });
    return true;
  }

  #create(key, role) {
    const state = windowStateKeeper({
      defaultWidth: role === ROLE_MEETING ? 1100 : 720,
      defaultHeight: role === ROLE_MEETING ? 720 : 760,
      file: `window-state-${role}.json`,
    });
    const offset = this.#cascade * CASCADE_STEP;
    this.#cascade = (this.#cascade + 1) % 8;

    const win = new BrowserWindow({
      title: "Teams for Linux",
      x: state.x === undefined ? undefined : state.x + offset,
      y: state.y === undefined ? undefined : state.y + offset,
      width: state.width,
      height: state.height,
      minWidth: MIN_WIDTH,
      minHeight: MIN_HEIGHT,
      backgroundColor: nativeTheme.shouldUseDarkColors ? "#302a75" : "#fff",
      autoHideMenuBar: this.#config.menubar === "auto",
      icon: this.#iconImage,
      frame: this.#config.frame,
      webPreferences: {
        partition: this.#config.partition,
        preload: path.join(__dirname, "..", "browser", "preload.js"),
        additionalArguments: [`${ROLE_ARGUMENT}${role}`],
        plugins: true,
        spellcheck: true,
        webviewTag: false,
        // Same settings as the main window: the preload needs DOM access to
        // the Teams React tree, and IPC is validated by security/ipcValidator.
        contextIsolation: false,
        nodeIntegration: false,
        sandbox: false,
      },
    });
    win.__teamsRole = role;
    state.manage(win);

    win.webContents.setWindowOpenHandler((details) => this.#handleWindowOpen(details));
    win.on("closed", () => {
      this.#windows.delete(key);
    });
    this.#windows.set(key, win);
    console.info("[POPOUT] Window opened", { role, open: this.#windows.size });
    return win;
  }
}

module.exports = {
  PopoutWindowManager,
  ROLE_ARGUMENT,
  buildChatUrl,
  isEnabled,
  isMeetingJoinUrl,
  meetingTarget,
  roleFromArgv,
};
