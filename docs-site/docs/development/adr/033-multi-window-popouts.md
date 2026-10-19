---
id: 033-multi-window-popouts
---

# ADR 033: Chat Pop-outs and Meeting Windows

## Status

✅ Implemented (2026-10-18). Generalises [ADR-032](032-secondary-window-for-calls.md) and supersedes [ADR-010](010-multiple-windows-support.md). Where this record and ADR-032 differ, this record decides.

## Context

The Windows and macOS clients open chats, channel conversations, apps and calendar views in their own windows that show only that content, and always give a meeting its own window. Teams for Linux had one window, so reading a thread while presenting, or keeping a meeting on a second monitor, was not possible. ADR-032 proposed a single opt-in secondary window for calls. The request grew to match the official client: pop out any view, show only its content, start meetings in a dedicated window, and work in Flatpak.

Spikes on 2026-10-18 against a signed-in personal account (`teams.live.com`) established the following about the Teams web client:

- **Teams web does not offer a pop-out.** The chat menus have no "open in new window" and nothing calls `window.open` for a chat. The code for it is in the bundle, but its feature flags (`enableChatsPopout`, `enableChatPopoutInContextMenu`) are set only for the desktop platforms ("Maglev", ids 48 to 51) and not for the browser platform (1415).
- **The platform is derived from the user agent, and so is the host mode.** A user agent containing `Teams/<a.b.c.d>` or `maglev/...` selects Maglev. That also puts Teams into native-host mode (`hostCommunication.loadModule("core")`, windows created by the host). With no host the page never boots: spoofing the user agent gives a blank window and `ExperienceLoader: Could not get loading window`. The flags cannot be enabled without the host mode, and user-level overrides come from a server-side config service, so this route is closed. Do not retry it.
- **A window is a secondary window when it has an opener on the same host.** In Teams' window provider, `opener !== null && opener.location.host === location.host` sets `viewportContext = "secondary"`. That already drops the top search bar and the "Chat |" title prefix.
- **A chat resolves in a secondary window, a meeting join does not, and the meeting window spins forever.** Tracing an opener-linked window that loads a `/meet/` link (every message to the opener, `BroadcastChannel`, history and network) showed the following:
  - The child boots as a secondary window and exchanges lifecycle events (`windowReady`, then `windowAppStart`, `coreInitialized`, `experienceStarted` are the names in the bundle) over a `multi-window-manager-channel`. That fabric assumes the opener's own Teams instance created the window through its window manager: on desktop hosts the host does, and the web build opens `about:blank?window=loading` and writes its HTML template into it. Here the opener is our code, so nothing on the opener side answers the handshake.
  - The child initialises as a main-window experience (it writes `mainWindowNavHistory`), then **the SPA itself rewrites its URL to the bare `/v2/`**, discarding the `#/meet/...&launchType=web` route, and renders only the loading spinner.
  - Auth discovery in that child fails with `tdmService not found in regionGtms table`, the only error of note.
  - It is not a timeout and not a network failure: the window stays on `/v2/` indefinitely. I could not isolate which of these is the single trigger; the handshake with no opener-side counterpart is the likeliest.
  - A chat route resolves because `/v2/#/l/chat/...` is a route a secondary window handles in place. A `/meet/` link is a deep link that Teams expects the primary window to handle.
  Meetings therefore use a window with no opener, which Teams boots as a normal primary window.
- **Every view renders in one layout with stable regions** (`experience-layout`): `app-layout-area--nav` (app bar), `app-layout-area--mid-nav` (list pane) and `app-layout-area--main` (content). Hiding the first two shows only the content, for a chat and for the calendar alike.

## Decision

Add a generic, **on by default** multi-window feature. The app opens the windows itself; it does not depend on Teams offering them.

- **`app/popoutWindows/`** owns the secondary windows: a registry keyed by chat thread or meeting URL (a repeat request focuses the existing window), per-role persisted bounds and a cascade offset.
- **Chat pop-outs are opened with `window.open`** from the renderer, with a frame name `tfl-popout:chat:<thread>`. The main window's `setWindowOpenHandler` allows exactly that shape (the URL must equal the launcher route rebuilt from the thread id, on a Teams host), so the new window has an opener and inherits the session partition. The manager adopts it on `did-create-window`.
- **Meeting joins** (`/l/meetup-join/`, `/meet/`) open in a `meeting` window created with `loadURL` and no opener, on three paths: a `window.open` from Teams when `urlHandling.openMeetupJoinInApp` is on, a deep link, and a second instance started with a meeting URL (the usual way a meeting link reaches the app on Linux). The last two no longer navigate the main window, so a call it is in is not disturbed, and a repeat link focuses the open meeting window. Chat and channel launchers matched by the same regex keep loading in place.
- **Focus layout.** In pop-out and meeting windows, `app/browser/tools/popoutLayout.js` injects a stylesheet that hides the app bar (`app-layout-area--nav`), the list pane (`--mid-nav`) and the global header (`--title-bar`, kept for frameless windows, where it carries the window controls). The grid tracks of the hidden regions collapse, so the header and content fill the window in both layouts. Positioning the content area by hand was tried and breaks one layout or the other (the pop-out's chat header lives in a sibling region). Every rule is scoped with `:has()` on the content area, so if Teams renames the regions the window is the full client instead of a blank window.
- **Window roles over IPC.** The preload asks the main process (`get-window-role`, by `webContents` id) for `main`, `popout` or `meeting`. Options passed to a `window.open` child cannot carry a launch argument. Pop-outs skip the app-wide modules (`settings`, `trayIconRenderer`, `mqttStatusMonitor`, `meetingStartDetector`, `dockIconRenderer`, `customStickers`), the `ActivityManager` (so a call rings once), and return a notification stub (so a message notifies once). Media tools, `speakingIndicator` and the call-state bridge stay.
- **Joining from inside Teams.** Teams handles the calendar Join button in place: no `window.open`, no URL change, no history entry (this was already established in ADR-032). Two things are visible, and neither depends on the interface language. The pre-join screen is rendered with stable ids (`calling-prejoin-screen`, `meeting-header-title`, `prejoin-cancel-button`, `prejoin-join-button`), and Teams writes a record of the join to `localStorage` under a key ending in `.react-web-client.get-user-media-data-bag`: a JSON string holding JSON with `joinOptions.context` (`calendar_grid_event_card_join_button` for the calendar), `joinOptions.title`, `joinOptions.showPrejoin` and `coordinates.meetingData.meetingUrl`. `app/browser/tools/meetingJoinRedirect.js` runs in the main window only. When the pre-join screen appears it reads that record, checks that the record's title is the one the screen shows (so a stale record never opens the wrong meeting), asks the main process over `popout-open-meeting` to open the link in a meeting window, and when one took it, waits for the pre-join to settle and clicks Teams' own Cancel. It reacts to the screen and not to the record being written, because a second Join for the same meeting makes Teams show the old pre-join again without writing a record. The main-process side validates the link (Teams host, join path) and honours `multiWindow.enabled` and `openMeetingsIn`.
- **Gestures.** Double-clicking a chat in the list (the official gesture) or the Window menu entry "Open Chat in New Window" (`Ctrl+Alt+O`, via `popout-open-current-chat`). The thread id is validated against an allowlist pattern.
- **Lifecycle.** Closing a secondary window closes only that window. It never hides to the tray and never quits the app. Closing the main window closes all of them and quits.
- **Call state is per renderer.** `call-connected` and `call-disconnected` update a set of calling senders. Screen-lock inhibition and the app-wide `teams-call-connected` / `teams-call-disconnected` events fire on the first call starting and the last call ending.
- **Configuration.** `multiWindow.enabled` (default `true`), `multiWindow.openMeetingsIn` (`window` or `main`, default `window`) and `multiWindow.focusLayout` (default `true`). `enabled: false` restores the single-window behaviour exactly.

Out of scope: moving a live call between windows ([#2499](https://github.com/IsmaelMartinez/teams-for-linux/issues/2499)), redirecting the Join button, "Add shared display" ([#2983](https://github.com/IsmaelMartinez/teams-for-linux/issues/2983)).

## Consequences

### Positive

- A chat pop-out matches the official client: the title is the chat name and the window shows the chat header, the messages and the compose box.
- Works anywhere the main window works. Windows are ordinary Chromium windows in the same process, so Flatpak needs no sandbox change.
- Calls in two windows are tracked correctly.

### Negative

- The layout depends on Microsoft's `data-tid` region names and on the `chat-list-item` / `chat-header` element ids. If Teams renames them, the gesture or the reduced layout stops working. The `:has()` scoping makes the layout fail safe; the gesture fails silently.
- The connected in-call screen of a meeting window has not been seen: the test environment blocks WebRTC media, so the join stops at "your network configuration is preventing us from making your call". The pre-join and join screens use the focus layout and look right; the in-call layout is unverified.
- A meeting joined from inside Teams (the calendar Join button, a meeting card, "Join with an ID") is redirected by watching the main window rather than by an API: Teams opens the pre-join screen in place and the module cancels it once the meeting window has the link (see Decision). The main window briefly shows the pre-join before it is cancelled, and Teams keeps its window title as "Calendar | Meeting join | <subject>" until the next navigation. A join with no pre-join screen (the user turned it off) starts the call straight away and cannot be cancelled cleanly, so it stays in the main window. Only consumer meeting links were seen: the join record for a work or school meeting may carry no link, in which case nothing is redirected.
- Calls (as opposed to meetings) started or answered in the main window stay there; moving a live call is out of scope.
- Channel conversations, apps and the calendar are not yet popped out, although the focus layout was verified on the calendar.
- Extra renderers add memory per open window, and meeting windows get no tray, badge or notifications of their own by design.

### Flatpak and Wayland

No sandbox changes. Window position restore and always-on-top are best effort on Wayland. Screen sharing from a meeting window uses the same per-session display-media handler as the main window; on Wayland each share still needs the portal picker. Not yet tested in a Flatpak build.

## Alternatives Considered

- **Make Teams render its own pop-out** (enable the desktop flags). Ruled out: the flags are tied to the desktop host mode, and spoofing it blanks the app.
- **Patch or intercept Microsoft's JavaScript to flip the flags.** Rejected as fragile and unsupported.
- **One fixed secondary window (ADR-032).** Does not cover chats or several meetings.
- **Opt-in rather than default.** Rejected: the official client behaves this way, and `multiWindow.enabled: false` is a one-line opt-out.

## Related

- [ADR-010](010-multiple-windows-support.md), [ADR-032](032-secondary-window-for-calls.md), [ADR-020](020-multi-account-profile-switcher.md)
- [#2971](https://github.com/IsmaelMartinez/teams-for-linux/issues/2971), [#1984](https://github.com/IsmaelMartinez/teams-for-linux/issues/1984)
- `app/popoutWindows/`, `app/browser/tools/popoutChat.js`, `app/browser/preload.js`
