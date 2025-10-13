# Corner TaskView

Electron tray helper for Windows that triggers **Win + Tab** (Task View) when your cursor touches a configurable screen corner.

## Prerequisites

- [Node.js](https://nodejs.org/) 18 or newer
- Windows 11

## Development setup

```powershell
npm install
npm start
```

The app does not open a window. A tray icon appears instead. Right-click the icon to choose:
- the corner to watch
- the target display
- the detection speed (Instant, Very fast, Fast, Medium, Slow)

The detection speed controls both the pointer polling frequency and the cooldown before the next Win + Tab trigger. The “Instant” preset runs an immediate loop for the lowest latency. The shortcut fires only once per entry into the corner until the pointer leaves it.

## Configuration

- Tray icon: replace `icon.png` at the project root to ship a custom icon.
- Corner threshold: tweak the `EDGE_THRESHOLD_PX` constant in `src/main.js` to widen or tighten the hot corner area (default 30 pixels).
- Preferences are stored via `electron-store` in the user profile (`corner`, `displayId`, `detectionSpeed`).

## Build a Windows installer (NSIS)

The project is configured for [electron-builder](https://www.electron.build/).

```powershell
npm run dist
```

The installer is generated under `dist/`. It uses NSIS with:
- directory selection (not one-click)
- optional desktop shortcut
- optional “launch on Windows startup” toggle (writes to HKCU Run)

The uninstaller cleans up the desktop shortcut and startup entry that the installer created.
