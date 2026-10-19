/**
 * Open a meeting in its own window when it is joined from inside Teams.
 *
 * The official client gives every meeting its own window, including one joined
 * from the calendar. Teams web handles that Join button in place: no
 * `window.open`, no URL change, no history entry, so nothing outside the page
 * sees it. What it does leave is the pre-join screen (`calling-prejoin-screen`)
 * and a record of the join in `localStorage`, under a key ending in
 * `.react-web-client.get-user-media-data-bag`. The value is a JSON string
 * holding a JSON object with:
 *
 * - `joinOptions.showPrejoin`             whether a pre-join screen is shown
 * - `joinOptions.title`                   the meeting title
 * - `joinOptions.context`                 where the join started, for example
 *                                         `calendar_grid_event_card_join_button`
 * - `coordinates.meetingData.meetingUrl`  the join link, passcode included
 *
 * None of it depends on the interface language. When the pre-join screen
 * appears in the main window, this module reads that record, checks that its
 * title is the one the screen shows, asks the main process to open the link in
 * a meeting window, and when that works cancels the in-place pre-join through
 * Teams' own Cancel button so the main window is left as it was.
 *
 * It reacts to the screen, not to the record being written, because a second
 * Join for the same meeting makes Teams show the old pre-join again without
 * writing a new record. A join with no pre-join screen starts the call
 * straight away and cannot be cancelled cleanly, so it is left alone. The link
 * and title are never logged: they identify a meeting.
 */

const { isMeetingJoinUrl } = require("../../popoutWindows/role");

const KEY_SUFFIX = ".react-web-client.get-user-media-data-bag";
const PREJOIN_SELECTOR = '[data-tid="calling-prejoin-screen"]';
const TITLE_SELECTOR = '[data-tid="meeting-header-title"]';
const CANCEL_SELECTOR = '[data-tid="prejoin-cancel-button"]';
const JOIN_SELECTOR = '[data-tid="prejoin-join-button"]';
const POLL_MS = 100;
const CANCEL_GIVE_UP_MS = 10000;
const TITLE_GIVE_UP_MS = 3000;
const DEBOUNCE_MS = 100;
// How long the pre-join must have been on screen before it is cancelled.
// Cancelling the moment it mounts leaves Teams' call state half torn down.
const SETTLE_MS = 800;

/**
 * Reads a join record as Teams stores it.
 *
 * @param {unknown} raw - the `localStorage` value
 * @returns {{url: string, title: string|null, showPrejoin: boolean}|null}
 */
function parseJoinRecord(raw) {
  if (typeof raw !== "string") {
    return null;
  }
  let value;
  try {
    value = JSON.parse(raw);
    // The record is stored as a JSON string that itself holds JSON.
    if (typeof value === "string") {
      value = JSON.parse(value);
    }
  } catch {
    return null;
  }
  const url = value?.coordinates?.meetingData?.meetingUrl;
  if (!isMeetingJoinUrl(url)) {
    return null;
  }
  const title = value.joinOptions?.title;
  return {
    url,
    title: typeof title === "string" && title.trim() ? title.trim() : null,
    showPrejoin: value.joinOptions?.showPrejoin === true,
  };
}

/**
 * @param {Storage} storage
 * @returns {{url: string, title: string|null, showPrejoin: boolean}|null} the
 *   join record Teams last wrote, if there is a usable one
 */
function readJoinRecord(storage) {
  for (let index = 0; index < storage.length; index += 1) {
    const key = storage.key(index);
    if (typeof key === "string" && key.endsWith(KEY_SUFFIX)) {
      const record = parseJoinRecord(storage.getItem(key));
      if (record) {
        return record;
      }
    }
  }
  return null;
}

/**
 * Resolves with the trimmed text of the first element matching `selector`, or
 * null when none has text before `giveUpMs`.
 *
 * @param {Document} doc
 * @param {string} selector
 * @param {number} giveUpMs
 * @returns {Promise<string|null>}
 */
function waitForText(doc, selector, giveUpMs) {
  return new Promise((resolve) => {
    const startedAt = Date.now();
    const timer = setInterval(() => {
      const text = doc.querySelector(selector)?.textContent?.trim();
      if (text) {
        clearInterval(timer);
        resolve(text);
      } else if (Date.now() - startedAt > giveUpMs) {
        clearInterval(timer);
        resolve(null);
      }
    }, POLL_MS);
  });
}

/**
 * Clicks Teams' own Cancel on the pre-join screen once it has settled.
 *
 * @param {Document} doc
 * @returns {Promise<boolean>} whether the pre-join was cancelled
 */
function cancelPrejoin(doc) {
  return new Promise((resolve) => {
    const startedAt = Date.now();
    let settledSince = null;
    const timer = setInterval(() => {
      const button = doc.querySelector(CANCEL_SELECTOR);
      if (button && doc.querySelector(JOIN_SELECTOR)) {
        settledSince ??= Date.now();
        if (Date.now() - settledSince >= SETTLE_MS) {
          clearInterval(timer);
          button.click();
          resolve(true);
          return;
        }
      } else {
        settledSince = null;
      }
      if (Date.now() - startedAt > CANCEL_GIVE_UP_MS) {
        clearInterval(timer);
        resolve(false);
      }
    }, POLL_MS);
  });
}

/**
 * Handles a pre-join screen that just appeared in the main window.
 *
 * @param {{doc: Document, storage: Storage, ipcRenderer: object}} env
 * @returns {Promise<boolean>} whether the meeting went to its own window
 */
async function redirectPrejoin({ doc, storage, ipcRenderer }) {
  const record = readJoinRecord(storage);
  if (!record?.showPrejoin) {
    return false;
  }
  // The record is the last join Teams wrote. Make sure it is the meeting this
  // screen shows, so a stale record never opens the wrong meeting.
  if (record.title) {
    const shown = await waitForText(doc, TITLE_SELECTOR, TITLE_GIVE_UP_MS);
    if (shown !== record.title) {
      return false;
    }
  }
  const opened = await ipcRenderer.invoke("popout-open-meeting", record.url);
  return opened ? cancelPrejoin(doc) : false;
}

function init(config, ipcRenderer) {
  if (config?.multiWindow?.enabled === false) {
    return;
  }

  let wasPresent = Boolean(document.querySelector(PREJOIN_SELECTOR));
  let timer = null;

  const check = () => {
    timer = null;
    const present = Boolean(document.querySelector(PREJOIN_SELECTOR));
    if (present && !wasPresent) {
      redirectPrejoin({ doc: document, storage: globalThis.localStorage, ipcRenderer }).catch(() => {});
    }
    wasPresent = present;
  };

  new MutationObserver(() => {
    timer ??= setTimeout(check, DEBOUNCE_MS);
  }).observe(document.documentElement, { childList: true, subtree: true });
}

module.exports = {
  init,
  parseJoinRecord,
  readJoinRecord,
  redirectPrejoin,
  cancelPrejoin,
  waitForText,
  KEY_SUFFIX,
};
