const { BrowserWindow, nativeTheme } = require("electron");
const path = require("node:path");
const windowStateKeeper = require("electron-window-state");
const {
  KIND_CHAT,
  ROLE_MAIN,
  ROLE_MEETING,
  ROLE_POPOUT,
  buildChatUrl,
  isMeetingJoinUrl,
  parsePopoutFrameName,
} = require("./role");

const MIN_WIDTH = 320;
const MIN_HEIGHT = 240;
const CASCADE_STEP = 30;

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
 * Owns the secondary Teams windows: chat pop-outs and meeting windows.
 *
 * A pop-out is opened by a renderer calling `window.open`, so the new window
 * has an opener on the same host. Teams treats such a window as a secondary
 * window, which is the same path its own desktop pop-outs take. The manager
 * decides in `handleWindowOpen` whether a `window.open` is a pop-out request,
 * hands Electron the window options, and adopts the created window to track
 * it, persist its bounds and install the window-open policy on it.
 *
 * Every window shares the session partition of the window that opened it, so
 * it is signed in without another prompt. A window is a peer that can be
 * closed on its own; it never hides to the tray and never quits the app. The
 * main window owns the tray, the badge and the app lifetime.
 */
class PopoutWindowManager {
  #config;
  #iconImage;
  #fallbackWindowOpen;
  #windows = new Map();
  #rolesByContentsId = new Map();
  #pending = new Map();
  #cascade = 0;

  /**
   * @param {object} deps
   * @param {object} deps.config - startup configuration
   * @param {(details: object) => object} deps.fallbackWindowOpen - window-open
   *   policy for everything that is not a pop-out (external links, auth popups)
   * @param {Electron.NativeImage} [deps.iconImage]
   */
  constructor({ config, fallbackWindowOpen, iconImage }) {
    this.#config = config;
    this.#fallbackWindowOpen = fallbackWindowOpen;
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
   * Decides what to do with a `window.open`. Returns null when the request is
   * not a pop-out, so the caller applies its normal policy.
   *
   * @param {{url: string, frameName?: string}} details
   * @returns {{action: "allow", overrideBrowserWindowOptions: object}|{action: "deny"}|null}
   */
  handleWindowOpen(details) {
    if (!isEnabled(this.#config)) {
      return null;
    }
    const request = this.#classify(details);
    if (!request) {
      return null;
    }
    const existing = this.#windows.get(request.key);
    if (existing && !existing.isDestroyed()) {
      this.#focus(existing);
      return { action: "deny" };
    }
    if (request.role === ROLE_MEETING) {
      this.#openMeetingWindow(request, details.url);
      return { action: "deny" };
    }
    this.#pending.set(this.#pendingKey(details), request);
    return {
      action: "allow",
      overrideBrowserWindowOptions: this.#windowOptions(request.role),
    };
  }

  /**
   * Opens a meeting join link in a meeting window, or focuses the window that
   * already has it. For links that reach the app from outside Teams (a deep
   * link, a second instance), which never pass through `window.open`.
   *
   * @param {string} url
   * @returns {boolean} true when a meeting window took the link
   */
  openMeeting(url) {
    if (!isEnabled(this.#config) || meetingTarget(this.#config) !== "window" || !isMeetingJoinUrl(url)) {
      return false;
    }
    const request = { key: `meeting:${url}`, role: ROLE_MEETING };
    const existing = this.#windows.get(request.key);
    if (existing && !existing.isDestroyed()) {
      this.#focus(existing);
    } else {
      this.#openMeetingWindow(request, url);
    }
    return true;
  }

  /**
   * Registers the window Electron created for an allowed request. Wire it to
   * `did-create-window` on every webContents that installs `handleWindowOpen`.
   *
   * @param {Electron.BrowserWindow} win
   * @param {{url: string, frameName?: string}} details
   */
  adopt(win, details) {
    const request = this.#pending.get(this.#pendingKey(details));
    if (!request) {
      return;
    }
    this.#pending.delete(this.#pendingKey(details));

    this.#register(win, request);
  }

  /**
   * The role of the window a renderer runs in. The preload asks for it over IPC
   * because options passed to a `window.open` child cannot carry a launch
   * argument, and a meeting window gets its frame name from Teams.
   *
   * @param {number} webContentsId
   * @returns {"main"|"popout"|"meeting"}
   */
  roleOf(webContentsId) {
    return this.#rolesByContentsId.get(webContentsId) ?? ROLE_MAIN;
  }

  closeAll() {
    for (const win of [...this.#windows.values()]) {
      if (!win.isDestroyed()) {
        win.destroy();
      }
    }
    this.#windows.clear();
    this.#rolesByContentsId.clear();
    this.#pending.clear();
  }

  /**
   * Installs the policy on a webContents that can open pop-outs: the main
   * window, and every window the manager creates.
   *
   * @param {Electron.WebContents} webContents
   */
  attach(webContents) {
    webContents.setWindowOpenHandler(
      (details) => this.handleWindowOpen(details) ?? this.#fallbackWindowOpen(details)
    );
    webContents.on("did-create-window", (win, details) => this.adopt(win, details));
  }

  // A meeting join link is loaded into a window of its own with no opener. Teams
  // boots a window that has an opener as a secondary window, and that mode does
  // not resolve a plain join link: it stays on the loading screen. A chat route
  // does resolve, so chats keep the opener (see handleWindowOpen).
  #openMeetingWindow(request, url) {
    const win = new BrowserWindow({
      ...this.#windowOptions(request.role),
      show: true,
      webPreferences: {
        partition: this.#config.partition,
        preload: path.join(__dirname, "..", "browser", "preload.js"),
        plugins: true,
        spellcheck: true,
        webviewTag: false,
        // Same settings as the main window.
        contextIsolation: false,
        nodeIntegration: false,
        sandbox: false,
      },
    });
    this.#register(win, request);
    // The SPA taking over the route aborts the first load in normal use, so a
    // rejected promise here is not a failure.
    win.loadURL(url, { userAgent: this.#config.chromeUserAgent }).catch((error) => {
      console.debug("[POPOUT] Navigation interrupted", { code: error?.code ?? error?.errno });
    });
  }

  #register(win, request) {
    const state = windowStateKeeper({
      defaultWidth: request.role === ROLE_MEETING ? 1100 : 720,
      defaultHeight: request.role === ROLE_MEETING ? 720 : 760,
      file: `window-state-${request.role}.json`,
    });
    state.manage(win);

    win.__teamsRole = request.role;
    const contentsId = win.webContents.id;
    this.#rolesByContentsId.set(contentsId, request.role);
    this.#windows.set(request.key, win);
    this.#installPolicy(win);
    win.on("closed", () => {
      this.#windows.delete(request.key);
      this.#rolesByContentsId.delete(contentsId);
    });
    console.info("[POPOUT] Window opened", { role: request.role, open: this.#windows.size });
  }

  #installPolicy(win) {
    this.attach(win.webContents);
  }

  #pendingKey(details) {
    return `${details.frameName ?? ""}|${details.url}`;
  }

  #classify(details) {
    const frame = parsePopoutFrameName(details.frameName);
    if (frame?.kind === KIND_CHAT) {
      let origin;
      try {
        origin = new URL(details.url).origin;
      } catch {
        return null;
      }
      // Rebuilding the URL from the thread id pins the request to the launcher
      // route: a renderer cannot use the pop-out name to open anything else.
      if (buildChatUrl(origin, frame.key) !== details.url) {
        console.warn("[POPOUT] Rejected chat pop-out request");
        return null;
      }
      return { key: `chat:${frame.key}`, role: ROLE_POPOUT };
    }
    // Meeting links open in a window only where they would otherwise open in
    // the app at all; with the in-app option off they keep being dropped.
    if (
      this.#config.onNewWindowOpenMeetupJoinUrlInApp &&
      meetingTarget(this.#config) === "window" &&
      isMeetingJoinUrl(details.url)
    ) {
      return { key: `meeting:${details.url}`, role: ROLE_MEETING };
    }
    return null;
  }

  #focus(win) {
    if (win.isMinimized()) {
      win.restore();
    }
    win.show();
    win.focus();
  }

  #windowOptions(role) {
    const state = windowStateKeeper({
      defaultWidth: role === ROLE_MEETING ? 1100 : 720,
      defaultHeight: role === ROLE_MEETING ? 720 : 760,
      file: `window-state-${role}.json`,
    });
    const offset = this.#cascade * CASCADE_STEP;
    this.#cascade = (this.#cascade + 1) % 8;
    return {
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
        // Same settings as the main window: the preload needs DOM access to
        // the Teams React tree, and IPC is validated by security/ipcValidator.
        preload: path.join(__dirname, "..", "browser", "preload.js"),
        contextIsolation: false,
        nodeIntegration: false,
        sandbox: false,
        webviewTag: false,
      },
    };
  }
}

module.exports = {
  PopoutWindowManager,
  buildChatUrl,
  isEnabled,
  isMeetingJoinUrl,
  meetingTarget,
};
