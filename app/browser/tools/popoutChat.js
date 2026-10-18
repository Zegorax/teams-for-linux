/**
 * Pop a chat out into its own window, as the official client does.
 *
 * Teams web does not offer a pop-out for chats, so this module opens the
 * chat's launcher route with `window.open`. The new window keeps an opener on
 * the same host, which makes Teams treat it as a secondary window, and the main
 * process recognises the frame name (see popoutWindows/) and creates it with
 * the pop-out role. Two gestures:
 *
 * - Double-clicking a chat in the list, which is the official client's gesture.
 * - The "Open Chat in New Window" menu item, which makes the main process send
 *   `popout-open-current-chat` to the focused window.
 *
 * The thread id is read from the element ids Teams renders for each chat
 * (`title-chat-list-item_<thread>` in the list, `chat-header-<thread>` in the
 * open chat). Nothing is logged: thread ids identify conversations.
 */

const { KIND_CHAT, buildChatUrl, popoutFrameName } = require("../../popoutWindows/role");

const LIST_ITEM_TITLE = '[id^="title-chat-list-item_"]';
const LIST_ITEM_PREFIX = "title-chat-list-item_";
const HEADER_SELECTOR = '[id^="chat-header-"]';
const HEADER_PREFIX = "chat-header-";

/**
 * Finds the thread id of the chat list item that contains `target`: the
 * nearest ancestor holding exactly one list-item title. An ancestor holding
 * several means the click landed outside any single item.
 *
 * @param {Element|null} target
 * @returns {string|null}
 */
function threadIdFromListTarget(target) {
  let element = target;
  while (element && element.nodeType === 1) {
    if (element.id?.startsWith(LIST_ITEM_PREFIX)) {
      return element.id.slice(LIST_ITEM_PREFIX.length) || null;
    }
    const titles = element.querySelectorAll(LIST_ITEM_TITLE);
    if (titles.length === 1) {
      return titles[0].id.slice(LIST_ITEM_PREFIX.length) || null;
    }
    if (titles.length > 1) {
      return null;
    }
    element = element.parentElement;
  }
  return null;
}

/**
 * @param {Document} doc
 * @returns {string|null} thread id of the chat open in the main pane
 */
function currentThreadId(doc) {
  const header = doc.querySelector(HEADER_SELECTOR);
  return header?.id.slice(HEADER_PREFIX.length) || null;
}

function openChat(threadId) {
  const url = buildChatUrl(globalThis.location.origin, threadId);
  if (url) {
    globalThis.open(url, popoutFrameName(KIND_CHAT, threadId));
  }
}

function init(config, ipcRenderer) {
  if (config?.multiWindow?.enabled === false) {
    return;
  }

  document.addEventListener("dblclick", (event) => {
    const threadId = threadIdFromListTarget(event.target);
    if (threadId) {
      openChat(threadId);
    }
  });

  ipcRenderer.on("popout-open-current-chat", () => {
    const threadId = currentThreadId(document);
    if (threadId) {
      openChat(threadId);
    }
  });
}

module.exports = { init, threadIdFromListTarget, currentThreadId };
