---
id: 033-multi-window-popouts
---

# ADR 033: Chat Pop-outs and Meeting Windows

## Status

✅ Implemented (2026-10-18). Generalises [ADR-032](032-secondary-window-for-calls.md) and supersedes [ADR-010](010-multiple-windows-support.md). Where this record and ADR-032 differ, this record decides.

## Context

The Windows and macOS clients open chats, channel conversations and apps in their own windows, and always give a meeting its own window. Teams for Linux had one window, so reading a thread while presenting, or keeping a meeting on a second monitor, was not possible. ADR-032 proposed a single opt-in secondary window for calls. The request grew to match the official client: pop out any chat, and start meetings in a dedicated window, working in Flatpak.

A spike on 2026-10-18 against a signed-in personal account (`teams.live.com`) established three facts:

- **Teams web does not ask for a pop-out.** Neither the chat list item menu nor the chat header menu offers "open in new window", and nothing calls `window.open` for a chat. The official client's pop-out is a desktop-host feature, as ADR-032 found for "Add shared display". The app therefore cannot wait for Teams to request a window.
- **A chat can be opened by URL in a fresh document.** The `/l/chat/<thread>/conversations` launcher route, already used by `deepLinkRouter.js`, resolves in page. The thread id is in the DOM as `title-chat-list-item_<thread>` (list) and `chat-header-<thread>` (open chat).
- **A second window on the same partition is signed in and independent**, including for a meeting join page.

## Decision

Add a generic, **on by default** multi-window feature. The app opens the windows itself; it does not depend on Teams.

- **`app/popoutWindows/`** owns the secondary windows: a registry keyed by chat thread or meeting URL (a repeat request focuses the existing window), per-role persisted bounds, and a cascade offset. Each window uses the main window's partition and `webPreferences`, so it is signed in with no extra prompt.
- **Window roles.** The main process passes `--teams-window-role=popout|meeting` to the preload. Pop-outs skip the app-wide modules (`settings`, `trayIconRenderer`, `mqttStatusMonitor`, `meetingStartDetector`, `dockIconRenderer`, `customStickers`), the `ActivityManager` (so a call rings once), and return a notification stub (so a message notifies once). Media tools, `speakingIndicator` and the call-state bridge stay.
- **Chats.** Double-clicking a chat in the list (the official gesture) or the Window menu entry "Open Chat in New Window" (`Ctrl+Alt+O`) asks the main process over `popout-open-chat`. The thread id is validated against an allowlist pattern and the origin must be a Teams host.
- **Meetings.** Meeting join links (`/l/meetup-join/`, `/meet/`) that previously loaded in the main window open a `meeting` window when `urlHandling.openMeetupJoinInApp` is on. Chat and channel launchers that the same regex matches keep loading in place.
- **Lifecycle.** Closing a secondary window closes only that window. It never hides to the tray and never quits the app. Closing the main window closes all of them and quits.
- **Call state is per renderer.** `call-connected` and `call-disconnected` update a set of calling senders. Screen-lock inhibition and the app-wide `teams-call-connected` / `teams-call-disconnected` events fire on the first call starting and the last call ending, so a second call never overwrites the first and MQTT never reports "not in a call" while one remains.
- **Configuration.** `multiWindow.enabled` (default `true`) and `multiWindow.openMeetingsIn` (`window` or `main`, default `window`). `enabled: false` restores the single-window behaviour exactly.

Out of scope: moving a live call between windows ([#2499](https://github.com/IsmaelMartinez/teams-for-linux/issues/2499)), redirecting the Join button, "Add shared display" ([#2983](https://github.com/IsmaelMartinez/teams-for-linux/issues/2983)), and a chat-only layout.

## Consequences

### Positive

- Matches the official client for chats and meetings without hooking the Teams UI beyond reading element ids.
- Works anywhere the main window works. Windows are ordinary Chromium windows in the same process, so Flatpak needs no sandbox change.
- Calls in two windows are tracked correctly.

### Negative

- **The pop-out is the full Teams client opened on the chat**, not the official chat-only pane. A chat-only layout would need CSS against Teams' markup and is deferred.
- Extra renderers add memory per open window.
- The double-click gesture and the ids it reads depend on Teams' DOM. If Teams renames them the gesture stops working silently; the menu entry shares the same lookup. Unit tests pin the lookup logic, not Teams' markup.
- Meeting windows get no tray, badge or notifications of their own by design.
- Channel conversations and apps are not yet popped out. They need their own routes and remain follow-ups.

### Flatpak and Wayland

No sandbox changes. Window position restore and always-on-top are best effort on Wayland. Screen sharing from a meeting window uses the same per-session display-media handler as the main window; on Wayland each share still needs the portal picker.

## Alternatives Considered

- **Wait for Teams to call `window.open`.** Ruled out by the spike: Teams web does not.
- **One fixed secondary window (ADR-032).** Does not cover chats or several meetings.
- **Opt-in rather than default.** Rejected: the official client behaves this way, and `multiWindow.enabled: false` is a one-line opt-out.

## Related

- [ADR-010](010-multiple-windows-support.md), [ADR-032](032-secondary-window-for-calls.md), [ADR-020](020-multi-account-profile-switcher.md)
- [#2971](https://github.com/IsmaelMartinez/teams-for-linux/issues/2971), [#1984](https://github.com/IsmaelMartinez/teams-for-linux/issues/1984)
- `app/popoutWindows/`, `app/browser/tools/popoutChat.js`, `app/browser/preload.js`
