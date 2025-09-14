// Windows-only module compiled via cfg in the binary's main.rs

use anyhow::Result;
use std::cell::RefCell;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    GetMonitorInfoW, MONITOR_DEFAULTTONEAREST, MONITORINFO, MonitorFromWindow,
};
use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::RemoteDesktop::{
    NOTIFY_FOR_THIS_SESSION, WTSRegisterSessionNotification, WTSUnRegisterSessionNotification,
};
use windows::Win32::UI::WindowsAndMessaging::*;

use mddskmgr::autorun;
use mddskmgr::config::{self, Config, Paths};
use mddskmgr::hotkeys::{self, HK_EDIT_DESC, HK_EDIT_TITLE, HK_TOGGLE};
use mddskmgr::overlay::Overlay;
use mddskmgr::tray;
use mddskmgr::tray::{
    CMD_EDIT_DESC, CMD_EDIT_TITLE, CMD_EXIT, CMD_OPEN_CONFIG, CMD_TOGGLE, TRAY_MSG, Tray,
};
use mddskmgr::ui;
use mddskmgr::vd;
use notify::{RecommendedWatcher, RecursiveMode, Watcher};
use std::sync::mpsc as std_mpsc;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::core::PCWSTR;

const WM_VD_SWITCHED: u32 = WM_APP + 2;
const WM_CFG_CHANGED: u32 = WM_APP + 3;

thread_local! {
    static APP: RefCell<Option<AppState>> = const { RefCell::new(None) };
}

struct AppState {
    hwnd: HWND,
    cfg: Config,
    cfg_paths: Paths,
    overlay: Overlay,
    current_guid: String,
    visible: bool,
    tray: Tray,
    taskbar_created_msg: u32,
    vd_thread: Option<winvd::DesktopEventThread>,
    hide_for_accessibility: bool,
    hide_for_fullscreen: bool,
    anchor_index: u8, // 0=1/4,1=1/2,2=3/4
}

fn compute_line(cfg: &Config, guid: &str) -> (String, i32) {
    let label = cfg.desktops.get(guid).cloned().unwrap_or_default();
    let title = if label.title.trim().is_empty() {
        "Desktop".to_string()
    } else {
        label.title
    };
    let desc = label.description;
    let line = format!("{} : {}", title, desc);
    (line, cfg.appearance.margin_px)
}

fn anchor_ratio_from_index(idx: u8) -> f32 {
    match idx % 3 {
        0 => 0.25,
        1 => 0.5,
        _ => 0.75,
    }
}

fn draw_overlay_line(overlay: &Overlay, cfg: &Config, guid: &str) {
    let (line, margin) = compute_line(cfg, guid);
    let hints = "(Ctrl+Alt+T,D,O,L)";
    tracing::debug!(guid=%guid, line=%line, "update_overlay_text");
    let ratio = APP.with(|slot| {
        if let Some(app) = &*slot.borrow() {
            anchor_ratio_from_index(app.anchor_index)
        } else {
            0.5
        }
    });
    let _ = overlay.draw_line_top_anchor_with_hints(&line, hints, margin, ratio);
}

fn is_high_contrast() -> bool {
    unsafe {
        let mut hc = windows::Win32::UI::Accessibility::HIGHCONTRASTW {
            cbSize: std::mem::size_of::<windows::Win32::UI::Accessibility::HIGHCONTRASTW>() as u32,
            ..Default::default()
        };
        if SystemParametersInfoW(
            SPI_GETHIGHCONTRAST,
            hc.cbSize,
            Some(&mut hc as *mut _ as *mut _),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
        .is_ok()
        {
            (hc.dwFlags & windows::Win32::UI::Accessibility::HCF_HIGHCONTRASTON)
                != windows::Win32::UI::Accessibility::HIGHCONTRASTW_FLAGS(0)
        } else {
            false
        }
    }
}

fn is_foreground_fullscreen(app: &AppState) -> bool {
    unsafe {
        let fg = GetForegroundWindow();
        if fg.0.is_null() || fg == app.hwnd {
            return false;
        }
        let mut rc = RECT::default();
        if GetWindowRect(fg, &mut rc).is_err() {
            return false;
        }
        // Compare to monitor bounds (not work area) to avoid hiding on maximized windows.
        let mon = MonitorFromWindow(fg, MONITOR_DEFAULTTONEAREST);
        let mut mi = MONITORINFO {
            cbSize: std::mem::size_of::<MONITORINFO>() as u32,
            ..Default::default()
        };
        if !GetMonitorInfoW(mon, &mut mi).as_bool() {
            return false;
        }
        let m = mi.rcMonitor;
        let tol = 2;
        let covers_monitor = rc.left <= m.left + tol
            && rc.top <= m.top + tol
            && rc.right >= m.right - tol
            && rc.bottom >= m.bottom - tol;
        if !covers_monitor {
            return false;
        }
        // If it covers the monitor, treat as fullscreen only when there's no caption (likely borderless fullscreen)
        let style = GetWindowLongPtrW(fg, GWL_STYLE) as u32;
        let has_caption = (style & WS_CAPTION.0) != 0;
        let fullscreen = covers_monitor && !has_caption;
        if fullscreen {
            tracing::debug!(style=%format!("0x{style:08X}"), "fullscreen detected");
        }
        fullscreen
    }
}

fn refresh_visibility_now() {
    // Avoid holding RefCell borrows across ShowWindow (can re-enter wndproc).
    let args = APP.with(|slot| {
        if let Some(app) = &*slot.borrow() {
            let should_show = mddskmgr::core::should_show(
                app.visible,
                app.hide_for_accessibility,
                app.hide_for_fullscreen,
            );
            Some((app.hwnd, should_show))
        } else {
            None
        }
    });
    if let Some((hwnd, should_show)) = args {
        APP.with(|slot| {
            if let Some(app) = &*slot.borrow() {
                tracing::debug!(
                    visible=%app.visible,
                    hc_hide=%app.hide_for_accessibility,
                    fs_hide=%app.hide_for_fullscreen,
                    state=%(if should_show { "SHOW" } else { "HIDE" }),
                    "refresh_visibility_now"
                );
            }
        });
        unsafe {
            let _ = ShowWindow(hwnd, if should_show { SW_SHOW } else { SW_HIDE });
            let _ = SetWindowPos(
                hwnd,
                HWND_TOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
        }
    }
}

fn quick_edit(edit_title: bool) {
    // Snapshot state without holding a mutable borrow during the modal UI.
    let snapshot = APP.with(|slot| {
        if let Some(app) = &*slot.borrow() {
            let key = app.current_guid.clone();
            let label = app.cfg.desktops.get(&key).cloned().unwrap_or_default();
            let caption = if edit_title {
                "Edit Desktop Title"
            } else {
                "Edit Desktop Description"
            };
            let hint = if edit_title {
                "Change the title"
            } else {
                "Change the description"
            };
            let initial = if edit_title {
                label.title
            } else {
                label.description
            };
            Some((
                app.hwnd,
                key,
                caption.to_string(),
                hint.to_string(),
                initial,
            ))
        } else {
            None
        }
    });

    if let Some((hwnd, key, caption, hint, initial)) = snapshot {
        tracing::debug!(caption=%caption, guid=%key, initial=%initial, "quick_edit start");
        if let Some(newtext) = ui::prompt_text(hwnd, &caption, &hint, &initial) {
            tracing::debug!(text=%newtext, "quick_edit: new text");
            let mut snap: Option<(Overlay, Config, String)> = None;
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    let entry = app.cfg.desktops.entry(key).or_default();
                    if edit_title {
                        entry.title = newtext;
                    } else {
                        entry.description = newtext;
                    }
                    let _ = mddskmgr::config::save_atomic(&app.cfg, &app.cfg_paths);
                    tracing::debug!(?app.cfg_paths.cfg_file, "quick_edit: saved config");
                    snap = Some((
                        app.overlay.clone(),
                        app.cfg.clone(),
                        app.current_guid.clone(),
                    ));
                }
            });
            if let Some((ov, cfg_clone, gid)) = snap {
                draw_overlay_line(&ov, &cfg_clone, &gid);
                refresh_visibility_now();
            }
        }
    }
}

extern "system" fn wndproc(hwnd: HWND, msg: u32, w: WPARAM, l: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            APP.with(|slot| {
                let (mut cfg, paths) = config::load_or_default().expect("config load");
                if cfg.hotkeys.snap_position.key.eq_ignore_ascii_case("S") {
                    cfg.hotkeys.snap_position.key = "L".into();
                    let _ = config::save_atomic(&cfg, &paths);
                }
                let overlay = Overlay::new(hwnd, &cfg.appearance.font_family, cfg.appearance.font_size_dip).expect("overlay");
                let taskbar_created_msg = unsafe { RegisterWindowMessageW(PCWSTR(windows::core::w!("TaskbarCreated").as_wide().as_ptr())) };
                let tray = Tray::new(hwnd, "Desktop Labeler").expect("tray");

                // Register hotkeys (warn on duplicates)
                let hk = &cfg.hotkeys;
                if mddskmgr::hotkeys::has_duplicates(hk) {
                    // Show a friendly tray balloon (without holding a RefCell borrow).
                    let _ = mddskmgr::tray::Tray::balloon_for(hwnd, "Hotkeys", "Duplicate hotkeys detected; adjust labels.json");
                }
                let _ = hotkeys::register(hwnd, hk.edit_title.ctrl, hk.edit_title.alt, hk.edit_title.shift, &hk.edit_title.key, HK_EDIT_TITLE);
                let _ = hotkeys::register(hwnd, hk.edit_description.ctrl, hk.edit_description.alt, hk.edit_description.shift, &hk.edit_description.key, HK_EDIT_DESC);
                let _ = hotkeys::register(hwnd, hk.toggle_overlay.ctrl, hk.toggle_overlay.alt, hk.toggle_overlay.shift, &hk.toggle_overlay.key, HK_TOGGLE);
                let _ = hotkeys::register(hwnd, hk.snap_position.ctrl, hk.snap_position.alt, hk.snap_position.shift, &hk.snap_position.key, hotkeys::HK_SNAP);

                let current_guid = vd::get_current_desktop_guid();
                let vd_thread = mddskmgr::vd::start_vd_events(hwnd, WM_VD_SWITCHED);
                let app = AppState { hwnd, cfg, cfg_paths: paths, overlay, current_guid, visible: true, tray, taskbar_created_msg, vd_thread, hide_for_accessibility: false, hide_for_fullscreen: false, anchor_index: 1 };
                // Draw initial line before storing
                let ov = app.overlay.clone();
                let cfg_clone = app.cfg.clone();
                let gid = app.current_guid.clone();
                *slot.borrow_mut() = Some(app);
                draw_overlay_line(&ov, &cfg_clone, &gid);
                start_runtime_services(hwnd);
            });
            LRESULT(0)
        }
        msg if {
            let mut is_taskbar = false;
            APP.with(|slot| {
                if let Some(app) = &*slot.borrow() { is_taskbar = msg == app.taskbar_created_msg; }
            });
            is_taskbar
        } => {
            // Re-add the tray icon without keeping a RefCell borrow during Shell calls.
            let _ = mddskmgr::tray::Tray::re_add_for(hwnd);
            LRESULT(0)
        }
        WM_RBUTTONUP | WM_CONTEXTMENU => {
            let _ = mddskmgr::tray::Tray::show_popup_menu(hwnd);
            LRESULT(0)
        }
        WM_SETCURSOR => {
            unsafe {
                let _ = SetCursor(LoadCursorW(None, IDC_ARROW).unwrap_or_default());
            }
            LRESULT(1)
        }
        WM_VD_SWITCHED => {
            // Update current GUID, then draw outside of the borrow to avoid re-entrancy
            let mut snapshot: Option<(Overlay, Config, String)> = None;
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    let id = vd::get_current_desktop_guid();
                    if id != app.current_guid {
                        app.current_guid = id.clone();
                    }
                    snapshot = Some((app.overlay.clone(), app.cfg.clone(), app.current_guid.clone()));
                }
            });
            if let Some((ov, cfg_clone, gid)) = snapshot { draw_overlay_line(&ov, &cfg_clone, &gid); }
            LRESULT(0)
        }
        WM_CFG_CHANGED => {
            // Reload config and apply labels/hotkeys; show any balloon outside borrow.
            let mut need_balloon = false;
            let mut snapshot: Option<(Overlay, Config, String, HWND)> = None;
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    if let Ok((new_cfg, _)) = mddskmgr::config::load_or_default() {
                        app.cfg = new_cfg;
                        // Re-register hotkeys
                        mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_TITLE);
                        mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_DESC);
                        mddskmgr::hotkeys::unregister(app.hwnd, HK_TOGGLE);
                        mddskmgr::hotkeys::unregister(app.hwnd, hotkeys::HK_SNAP);
                        let hk = &app.cfg.hotkeys;
                        let ok1 = mddskmgr::hotkeys::register(app.hwnd, hk.edit_title.ctrl, hk.edit_title.alt, hk.edit_title.shift, &hk.edit_title.key, HK_EDIT_TITLE).unwrap_or(false);
                        let ok2 = mddskmgr::hotkeys::register(app.hwnd, hk.edit_description.ctrl, hk.edit_description.alt, hk.edit_description.shift, &hk.edit_description.key, HK_EDIT_DESC).unwrap_or(false);
                        let ok3 = mddskmgr::hotkeys::register(app.hwnd, hk.toggle_overlay.ctrl, hk.toggle_overlay.alt, hk.toggle_overlay.shift, &hk.toggle_overlay.key, HK_TOGGLE).unwrap_or(false);
                        let ok4 = mddskmgr::hotkeys::register(app.hwnd, hk.snap_position.ctrl, hk.snap_position.alt, hk.snap_position.shift, &hk.snap_position.key, hotkeys::HK_SNAP).unwrap_or(false);
                        if !(ok1 && ok2 && ok3 && ok4) { need_balloon = true; }
                        snapshot = Some((app.overlay.clone(), app.cfg.clone(), app.current_guid.clone(), app.hwnd));
                    }
                }
            });
            if let Some((ov, cfg_clone, gid, _)) = snapshot { draw_overlay_line(&ov, &cfg_clone, &gid); }
            if need_balloon {
                let _ = mddskmgr::tray::Tray::balloon_for(hwnd, "Hotkeys", "Some hotkeys failed to register. Adjust in labels.json");
            }
            LRESULT(0)
        }
        WM_TIMER => {
            if w.0 == 1 { // VD poller
                let mut snapshot: Option<(Overlay, Config, String)> = None;
                APP.with(|slot| {
                    if let Some(app) = &mut *slot.borrow_mut() {
                        let id = vd::get_current_desktop_guid();
                        if id != app.current_guid {
                            app.current_guid = id.clone();
                        }
                        snapshot = Some((app.overlay.clone(), app.cfg.clone(), app.current_guid.clone()));
                    }
                });
                if let Some((ov, cfg_clone, gid)) = snapshot { draw_overlay_line(&ov, &cfg_clone, &gid); }
            } else if w.0 == 2 {
                APP.with(|slot| {
                    if let Some(app) = &mut *slot.borrow_mut() {
                        let hide = if app.cfg.appearance.hide_on_fullscreen { is_foreground_fullscreen(app) } else { false };
                        app.hide_for_fullscreen = hide;
                    }
                });
            }
            if w.0 == 2 { refresh_visibility_now(); }
            if w.0 == 3 {
                // Keep overlay at the top of TOPMOST band without stealing focus
                let visible = APP.with(|slot| {
                    if let Some(app) = &*slot.borrow() {
                        mddskmgr::core::should_show(
                            app.visible,
                            app.hide_for_accessibility,
                            app.hide_for_fullscreen,
                        )
                    } else { false }
                });
                if visible { unsafe { let _ = SetWindowPos(hwnd, HWND_TOPMOST, 0,0,0,0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE); } }
            }
            LRESULT(0)
        }
        WM_SETTINGCHANGE => {
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    app.hide_for_accessibility = is_high_contrast();
                }
            });
            refresh_visibility_now();
            LRESULT(0)
        }
        0x02B1 /* WM_WTSSESSION_CHANGE */ => {
            let code = w.0 as u32;
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    match code { // 0x7 lock, 0x8 unlock
                        0x7 => { app.hide_for_accessibility = true; }
                        0x8 => { app.hide_for_accessibility = is_high_contrast(); }
                        _ => {}
                    }
                }
            });
            refresh_visibility_now();
            LRESULT(0)
        }
        WM_HOTKEY => {
            let id = w.0 as i32;
            let mut need_refresh = false;
            match id {
                HK_EDIT_TITLE => quick_edit(true),
                HK_EDIT_DESC => quick_edit(false),
                HK_TOGGLE => {
                    APP.with(|slot| {
                        if let Some(app) = &mut *slot.borrow_mut() { app.visible = !app.visible; }
                    });
                    need_refresh = true;
                }
                hotkeys::HK_SNAP => {
                    let mut snap: Option<(Overlay, Config, String, f32)> = None;
                    APP.with(|slot| {
                        if let Some(app) = &mut *slot.borrow_mut() {
                            app.anchor_index = (app.anchor_index + 1) % 3;
                            let ratio = anchor_ratio_from_index(app.anchor_index);
                            snap = Some((app.overlay.clone(), app.cfg.clone(), app.current_guid.clone(), ratio));
                        }
                    });
                    if let Some((ov, cfg_clone, gid, _ratio)) = snap { draw_overlay_line(&ov, &cfg_clone, &gid); }
                }
                _ => {}
            }
            if need_refresh { refresh_visibility_now(); }
            LRESULT(0)
        }
        TRAY_MSG => {
            let l = l.0 as u32;
            match l {
                WM_CONTEXTMENU | WM_RBUTTONUP => { let _ = mddskmgr::tray::Tray::show_popup_menu(hwnd); }
                WM_LBUTTONDBLCLK => {
                    APP.with(|slot| {
                        if let Some(app) = &mut *slot.borrow_mut() { app.visible = true; }
                    });
                    refresh_visibility_now();
                }
                _ => {}
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            unsafe { let _ = DestroyWindow(hwnd); }
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd = (w.0 & 0xFFFF) as u16;
            match cmd {
                CMD_EDIT_TITLE => quick_edit(true),
                CMD_EDIT_DESC => quick_edit(false),
                CMD_TOGGLE => {
                    APP.with(|slot| {
                        if let Some(app) = &mut *slot.borrow_mut() { app.visible = !app.visible; }
                    });
                    refresh_visibility_now();
                }
                CMD_OPEN_CONFIG => {
                    // Snapshot path then ShellExecute without holding borrow.
                    let path = APP.with(|slot| {
                        slot.borrow()
                            .as_ref()
                            .map(|app| app.cfg_paths.cfg_file.to_string_lossy().to_string())
                    });
                    if let Some(path) = path {
                        let wpath: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
                        unsafe { let _ = ShellExecuteW(None, PCWSTR(windows::core::w!("open").as_wide().as_ptr()), PCWSTR(wpath.as_ptr()), None, None, SW_SHOWNORMAL); }
                    }
                }
                CMD_EXIT => {
                    // Trigger orderly teardown to avoid hangs: destroy window -> WM_DESTROY posts quit.
                    unsafe { let _ = DestroyWindow(hwnd); }
                },
                tray::CMD_RUN_AT_STARTUP => {
                    let cur = autorun::get_run_at_login();
                    let _ = autorun::set_run_at_login(!cur);
                }
                tray::CMD_ABOUT => {
                    unsafe {
                        let _ = MessageBoxW(
                            hwnd,
                            PCWSTR(windows::core::w!("Desktop Labeler\r\n\r\nShows a per-desktop title overlay on the primary monitor.\r\n\r\nHotkeys:\r\n  Ctrl+Alt+T  Edit Title\r\n  Ctrl+Alt+D  Edit Description\r\n  Ctrl+Alt+O  Toggle Overlay\r\n  Ctrl+Alt+L  Snap Position").as_wide().as_ptr()),
                            PCWSTR(windows::core::w!("About Desktop Labeler").as_wide().as_ptr()),
                            MB_OK | MB_ICONINFORMATION,
                        );
                    }
                }
                _ => {}
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    // Stop timers to avoid re-entrancy during teardown
                    unsafe {
                        let _ = KillTimer(hwnd, 1);
                        let _ = KillTimer(hwnd, 2);
                        let _ = KillTimer(hwnd, 3);
                    }
                    mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_TITLE);
                    mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_DESC);
                    mddskmgr::hotkeys::unregister(app.hwnd, HK_TOGGLE);
                    mddskmgr::hotkeys::unregister(app.hwnd, hotkeys::HK_SNAP);
                    // Remove tray icon to prevent ghost icons after exit
                    app.tray.remove_icon();
                    // Drop virtual desktop event thread if present
                    app.vd_thread = None;
                }
            });
            unsafe { let _ = WTSUnRegisterSessionNotification(hwnd); }
            unsafe { PostQuitMessage(0); }
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, w, l) }
    }
}

fn single_instance_guard() -> bool {
    unsafe {
        let class_name = windows::core::w!("DesktopOverlayWndClass");
        let h = FindWindowW(class_name, None).unwrap_or(HWND(std::ptr::null_mut()));
        h.0.is_null()
    }
}

fn start_runtime_services(hwnd: HWND) {
    // Start VD watcher: prefer event thread; fall back to timer poller
    APP.with(|slot| {
        // First, immutable borrow for setup and to grab cfg_path
        let cfg_path_opt = {
            let borrowed = slot.borrow();
            if let Some(app) = &*borrowed {
                if app.vd_thread.is_none() {
                    unsafe {
                        SetTimer(hwnd, 1, 250, None);
                    }
                    vd::start_vd_poller(hwnd, WM_VD_SWITCHED);
                }
                unsafe {
                    SetTimer(hwnd, 2, 1000, None);
                }
                // Periodic topmost reassertion
                unsafe {
                    SetTimer(hwnd, 3, 1200, None);
                }
                unsafe {
                    let _ = WTSRegisterSessionNotification(hwnd, NOTIFY_FOR_THIS_SESSION);
                }
                Some(app.cfg_paths.cfg_file.clone())
            } else {
                None
            }
        };
        // Then, mutable borrow to set accessibility/visibility flags
        if let Some(app) = &mut *slot.borrow_mut() {
            app.hide_for_accessibility = is_high_contrast();
        }
        refresh_visibility_now();
        // Launch config watcher threads outside of any RefCell borrow
        if let Some(cfg_path) = cfg_path_opt {
            let (tx, rx) = std_mpsc::channel::<()>();
            std::thread::spawn(move || {
                let (watch_tx, watch_rx) = std_mpsc::channel();
                let mut watcher: RecommendedWatcher =
                    Watcher::new(watch_tx, notify::Config::default()).expect("watcher");
                let _ = watcher.watch(&cfg_path, RecursiveMode::NonRecursive);
                while let Ok(_ev) = watch_rx.recv() {
                    let _ = tx.send(());
                }
            });
            let hwnd_copy = hwnd.0 as usize;
            std::thread::spawn(move || {
                while rx.recv().is_ok() {
                    unsafe {
                        let _ = PostMessageW(
                            HWND(hwnd_copy as *mut std::ffi::c_void),
                            WM_CFG_CHANGED,
                            WPARAM(0),
                            LPARAM(0),
                        );
                    }
                }
            });
        }
    });
}

pub fn main() -> Result<()> {
    // Logging is initialized by src/main.rs; nothing to do here.

    if !single_instance_guard() {
        tracing::warn!("Another instance is already running. Exiting.");
        return Ok(());
    }

    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;

        let class_name = windows::core::w!("DesktopOverlayWndClass");
        let hinst = GetModuleHandleW(None).unwrap();
        let wc = WNDCLASSW {
            lpfnWndProc: Some(wndproc),
            hInstance: hinst.into(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            lpszClassName: class_name,
            ..Default::default()
        };
        RegisterClassW(&wc);

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE(
                (WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_TOPMOST | WS_EX_NOACTIVATE).0,
            ),
            class_name,
            windows::core::w!(""),
            WS_POPUP,
            0,
            0,
            400,
            40,
            None,
            None,
            hinst,
            None,
        )?;
        // Show first, then pin across desktops to avoid early 'WindowNotFound' logs in some shells
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = winvd::pin_window(hwnd);

        let mut msg = MSG::default();
        while GetMessageW(&mut msg, HWND(std::ptr::null_mut()), 0, 0).into() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        CoUninitialize();
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::panic::catch_unwind;

    #[test]
    fn start_runtime_no_panic() {
        // Build a minimal AppState with dummy tray to avoid Shell_NotifyIconW
        APP.with(|slot| {
            let (cfg, paths) = mddskmgr::config::load_or_default().expect("cfg");
            let overlay = mddskmgr::overlay::Overlay::new(
                HWND(std::ptr::null_mut()),
                &cfg.appearance.font_family,
                cfg.appearance.font_size_dip,
            )
            .unwrap();
            let tray = mddskmgr::tray::Tray {
                nid: unsafe { std::mem::zeroed() },
            };
            let app = super::AppState {
                hwnd: HWND(std::ptr::null_mut()),
                cfg,
                cfg_paths: paths,
                overlay,
                current_guid: "default".into(),
                visible: true,
                tray,
                taskbar_created_msg: 0,
                vd_thread: None,
                hide_for_accessibility: false,
                hide_for_fullscreen: false,
                anchor_index: 1,
            };
            *slot.borrow_mut() = Some(app);
        });

        let res = catch_unwind(|| {
            start_runtime_services(HWND(std::ptr::null_mut()));
        });
        assert!(res.is_ok());
    }

    // Smoke test: create a small window with a test wndproc that initializes
    // APP and calls start_runtime_services; ensure no panic and message loop runs.
    #[test]
    fn window_smoke_create() {
        unsafe extern "system" fn test_wndproc(
            hwnd: HWND,
            msg: u32,
            w: WPARAM,
            l: LPARAM,
        ) -> LRESULT {
            match msg {
                WM_CREATE => {
                    APP.with(|slot| {
                        let (cfg, paths) = mddskmgr::config::load_or_default().expect("cfg");
                        let overlay = mddskmgr::overlay::Overlay::new(
                            hwnd,
                            &cfg.appearance.font_family,
                            cfg.appearance.font_size_dip,
                        )
                        .unwrap();
                        let tray = mddskmgr::tray::Tray {
                            nid: unsafe { std::mem::zeroed() },
                        };
                        let app = AppState {
                            hwnd,
                            cfg,
                            cfg_paths: paths,
                            overlay,
                            current_guid: "default".into(),
                            visible: true,
                            tray,
                            taskbar_created_msg: 0,
                            vd_thread: None,
                            hide_for_accessibility: false,
                            hide_for_fullscreen: false,
                            anchor_index: 1,
                        };
                        *slot.borrow_mut() = Some(app);
                    });
                    start_runtime_services(hwnd);
                    LRESULT(0)
                }
                WM_DESTROY => {
                    unsafe {
                        PostQuitMessage(0);
                    }
                    LRESULT(0)
                }
                _ => unsafe { DefWindowProcW(hwnd, msg, w, l) },
            }
        }

        unsafe {
            let class_name = windows::core::w!("OverlayTestWndClass");
            let hinst = GetModuleHandleW(None).unwrap();
            let wc = WNDCLASSW {
                lpfnWndProc: Some(test_wndproc),
                hInstance: hinst.into(),
                hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
                lpszClassName: class_name,
                ..Default::default()
            };
            RegisterClassW(&wc);
            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                class_name,
                windows::core::w!(""),
                WS_POPUP,
                0,
                0,
                100,
                100,
                None,
                None,
                hinst,
                None,
            )
            .unwrap();
            let _ = ShowWindow(hwnd, SW_HIDE);
            // Pump a few messages then destroy
            let mut processed = 0u32;
            let mut msg = MSG::default();
            while processed < 10 {
                if PeekMessageW(&mut msg, HWND(std::ptr::null_mut()), 0, 0, PM_REMOVE).as_bool() {
                    let _ = TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                    processed += 1;
                } else {
                    // Post a destroy to exit
                    let _ = PostMessageW(hwnd, WM_DESTROY, WPARAM(0), LPARAM(0));
                }
            }
        }
    }
}
