# Desktop Labeler (mddsklbl)

A lightweight Windows overlay that shows a per–virtual-desktop title at the top center of your primary monitor. It is always-on-top, transparent, and pinned across desktops. Use global hotkeys or the tray menu to update the title/description for the current virtual desktop. Text is persisted per desktop and restored on restart.

## Features
- Per-desktop labels (Title and optional Description), persisted in JSON.
- Event-driven virtual desktop switch detection (Windows 11 24H2+), with seamless polling fallback.
- Transparent, crisp text rendering (DirectWrite/Direct2D) with a subtle backdrop; GDI fallback.
- Global hotkeys (configurable) and a small system tray menu.
- DPI-aware placement (Per-Monitor v2) and primary work-area centering.
- Accessibility- and focus-friendly behavior: auto-hide in High Contrast or when an app is fullscreen.

## Requirements
- Windows 11. Event notifications need Windows 11 24H2+; earlier builds fall back to a light poller.
- Rust toolchain (if building): stable Rust, `cargo`, MSVC target.

## Build & Run (from source)
```powershell
# In this repository
cargo build --release
# Run (dev)
cargo run
```
The first launch creates a default configuration file under your roaming profile.

## Tray Menu & Hotkeys
- Tray menu: Edit Title, Edit Description, Toggle Overlay, Open Config, Exit.
- Default hotkeys (changeable in config):
  - Ctrl+Alt+T — Edit Title
  - Ctrl+Alt+D — Edit Description
  - Ctrl+Alt+O — Toggle overlay visibility
  - Ctrl+Alt+L — Snap overlay position (cycle 1/4, 1/2, 3/4)
If any hotkey cannot be registered (OS conflict), it is skipped; adjust in the config.

## Configuration
Configuration is stored per-user at:
```
%APPDATA%\Acme\DesktopLabeler\config\labels.json
```
The app writes atomically (temp file + replace). A minimal schema:
```json
{
  "desktops": {
    "{GUID}": { "title": "Work", "description": "Focus on tickets" },
    "{GUID}": { "title": "Meetings", "description": "Teams/Zoom" }
  },
  "hotkeys": {
    "edit_title":       { "ctrl": true, "alt": true, "shift": false, "key": "T" },
    "edit_description": { "ctrl": true, "alt": true, "shift": false, "key": "D" },
    "toggle_overlay":   { "ctrl": true, "alt": true, "shift": false, "key": "O" },
    "snap_position":    { "ctrl": true, "alt": true, "shift": false, "key": "L" }
  },
  "appearance": {
    "font_family": "Segoe UI",
    "font_size_dip": 16,
    "margin_px": 8,
    "hide_on_fullscreen": false
  }
}
```
Notes
- Desktop keys are the OS GUIDs for each virtual desktop. The app discovers the current GUID automatically; you don’t need to prefill them.
- The edit dialogs enforce a simple input cap (200 chars) to keep the overlay tidy.

## Visibility & Accessibility
The overlay’s visibility is governed by:
- Your toggle state (hotkey or tray → Toggle Overlay)
- High Contrast mode: overlay auto-hides when OS High Contrast is ON; restores when OFF
- Fullscreen detection: hides if a foreground window fully covers the primary monitor
Together: the overlay shows only when Toggle=ON AND not High Contrast AND not Fullscreen.

## Virtual Desktop Detection
- Preferred: winvd event listener on Windows 11 24H2+ for instant switches.
- Fallback: a 250ms poller (low CPU) if events are unavailable.
- The overlay window is pinned to all desktops so it remains present; only the text changes with the current GUID.

## Rendering & Placement
- DirectWrite + Direct2D draw the label with per-pixel alpha onto a 32-bit top-down DIB, then `UpdateLayeredWindow` presents it.
- A subtle translucent backdrop improves legibility over busy wallpapers.
- Placement uses the primary monitor’s work area (excludes taskbar): centered horizontally, offset by `appearance.margin_px` from the top.

## Logging
Logs are written to `%LOCALAPPDATA%\Acme\DesktopLabeler\logs\mddsklbl.YYYY-MM-DD.log`. Control verbosity with `RUST_LOG` (e.g., `RUST_LOG=info` or `RUST_LOG=debug`).

## Troubleshooting
- Overlay not visible
  - High Contrast may be active — disable it, or rely on tray Toggle.
  - A fullscreen app may be in the foreground — Alt+Tab out or exit fullscreen.
  - You may have toggled it off — use Ctrl+Alt+O or the tray menu.
- Hotkey didn’t work
  - Some combinations are reserved by Windows; pick alternatives in `labels.json`.
- Titles don’t follow desktop switches
  - On older Windows 11 builds (pre-24H2), the app uses polling. It should still update within ~250ms.
- Multiple instances
  - The app runs as a single instance by class detection; if an instance is already running, a new one will exit.

## Development
- Format & lint: `cargo fmt --all` and `cargo clippy -- -D warnings`
- Tests: `cargo test`
- Run (verbose logs): set `RUST_LOG` and run with `cargo run`

Project layout
```
src/
  config.rs   # JSON schema + atomic save/load
  hotkeys.rs  # Register/Unregister helpers and IDs
  vd.rs       # Virtual desktop GUID + event/poller
  tray.rs     # Shell_NotifyIconW tray and menu
  overlay.rs  # Layered-window renderer (DWrite/D2D with fallback)
  ui.rs       # Minimal input dialog (Edit Title/Description)
  lib.rs      # Module exports
  main.rs     # Win32 window, message loop, wiring
```

## Security & Privacy
- No elevated privileges required; writes only under your `%APPDATA%/Acme/DesktopLabeler` directory.
- No network access or telemetry.

## Roadmap
- Config hot-reload via file watcher, refined logging, packaged installer (Inno/WiX), and optional Run‑at‑Login toggle.

---
If anything is unclear or you want this tailored for your deployment (MSI packaging, group policy defaults, etc.), open an issue or ask for a customized build profile.
