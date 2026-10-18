# Pop-out Windows

Chat pop-outs and meeting windows (ADR-033). Secondary windows run the Teams web app on the main window's session partition, so they are signed in already.

## Files

- `index.js` - `PopoutWindowManager` (window-open policy, registry, bounds, lifecycle) plus `isMeetingJoinUrl` and the config helpers. Main process only.
- `role.js` - window roles, the list of modules that run in the main window only, the pop-out frame-name protocol and `buildChatUrl`. No Electron imports, so the preload and the renderer modules can load it.

## How it fits together

- **Chats.** `app/browser/tools/popoutChat.js` reads the chat's thread id from the Teams DOM (double-click on a chat, or the "Open Chat in New Window" menu item) and calls `window.open(url, "tfl-popout:chat:<thread>")`. The main window's window-open handler is `manager.attach(webContents)`, which allows that shape and adopts the new window on `did-create-window`. The window keeps an opener on the same host, which Teams treats as a secondary window.
- **Meetings.** Meeting join links (`/l/meetup-join/`, `/meet/`) are loaded into a window with **no opener**, from a `window.open`, a deep link or a second instance (`manager.openMeeting`). An opener-linked meeting window spins forever: Teams boots it as a secondary window that waits for a handshake only the opener's own Teams instance can answer, and drops the deep link route (see ADR-033).
- **Roles.** The preload asks the main process for its role (`get-window-role`, by `webContents` id) and filters its module list, so a pop-out does not run tray, badge or MQTT code or raise a second copy of every notification.
- **Layout.** `app/browser/tools/popoutLayout.js` hides the app bar, the list pane and the global header in pop-out and meeting windows, and the grid collapses around the content. Rules are scoped with `:has()` on the content area, so a Teams rename leaves the full client. Do not position the content area by hand; it breaks one layout or the other.
- **Calls.** `BrowserWindowManager` tracks calls per renderer, so a call in a meeting window and one in the main window do not overwrite each other.

## Configuration

`multiWindow.enabled` (default `true`), `multiWindow.openMeetingsIn` (`window` | `main`) and `multiWindow.focusLayout` (default `true`). Setting `enabled` to `false` restores the single-window behaviour.

## Notes

- Never log thread ids or meeting URLs.
- Closing a secondary window never quits the app; closing the main window closes them all.
- Do not try to enable Teams' own pop-out by changing the user agent: it selects the desktop host mode and the app does not boot without a host.
