# Pop-out Windows

Chat pop-outs and meeting windows (ADR-033). Secondary windows run the Teams web app on the main window's session partition, so they are signed in already.

## Files

- `index.js` - `PopoutWindowManager` (registry, bounds, lifecycle) plus `buildChatUrl`, `isMeetingJoinUrl` and the config helpers. Main process only.
- `role.js` - window role parsing and the list of modules that run in the main window only. No Electron imports, so the preload can load it.

## How it fits together

- The renderer module `app/browser/tools/popoutChat.js` reads the chat's thread id from the Teams DOM and calls the `popout-open-chat` IPC channel (double-click on a chat, or the "Open Chat in New Window" menu item).
- `onNewWindow` in `app/mainAppWindow/index.js` sends meeting join links to `openMeeting` instead of loading them in the main window.
- The preload filters its module list by role, so a pop-out does not run tray, badge or MQTT code.
- `BrowserWindowManager` tracks calls per renderer, so a call in a meeting window and one in the main window do not overwrite each other.

## Configuration

`multiWindow.enabled` (default `true`) and `multiWindow.openMeetingsIn` (`window` | `main`). Setting `enabled` to `false` restores the single-window behaviour.

## Notes

- Never log thread ids or meeting URLs.
- Closing a secondary window never quits the app; closing the main window closes them all.
