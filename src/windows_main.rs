#![cfg(windows)]

use anyhow::Result;
use std::cell::RefCell;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::RemoteDesktop::{
    NOTIFY_FOR_THIS_SESSION, WTSRegisterSessionNotification, WTSUnRegisterSessionNotification,
};
use windows::Win32::UI::WindowsAndMessaging::*;

use mddskmgr::config::{self, Config, Paths};
use mddskmgr::hotkeys::{self, HK_EDIT_DESC, HK_EDIT_TITLE, HK_TOGGLE};
use mddskmgr::overlay::Overlay;
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
}

fn update_overlay_text(app: &mut AppState) {
    let label = app
        .cfg
        .desktops
        .get(&app.current_guid)
        .cloned()
        .unwrap_or_default();
    let title = if label.title.trim().is_empty() {
        "Desktop".to_string()
    } else {
        label.title
    };
    let desc = label.description;
    let line = format!("{}:{}", title, desc);
    let hints = "(Ctrl+Alt+T, Ctrl+Alt+D)";
    let margin = app.cfg.appearance.margin_px;
    eprintln!(
        "update_overlay_text: guid={}, title='{}', desc='{}' -> line='{}'",
        app.current_guid, title, desc, line
    );
    let _ = app
        .overlay
        .draw_line_top_center_with_hints(&line, hints, margin);
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
        // Compare to primary work area
        let mut work = RECT::default();
        let _ = SystemParametersInfoW(
            SPI_GETWORKAREA,
            0,
            Some(&mut work as *mut _ as *mut _),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        );
        let tol = 2; // small tolerance in pixels
        rc.left <= work.left + tol
            && rc.top <= work.top + tol
            && rc.right >= work.right - tol
            && rc.bottom >= work.bottom - tol
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
                eprintln!(
                    "refresh_visibility_now: visible={}, hc_hide={}, fs_hide={} => {}",
                    app.visible,
                    app.hide_for_accessibility,
                    app.hide_for_fullscreen,
                    if should_show { "SHOW" } else { "HIDE" }
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
        eprintln!(
            "quick_edit: '{}' for guid={}, initial='{}'",
            caption, key, initial
        );
        if let Some(newtext) = ui::prompt_text(hwnd, &caption, &hint, &initial) {
            eprintln!("quick_edit: new text='{}'", newtext);
            let mut updated = false;
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    let entry = app.cfg.desktops.entry(key).or_default();
                    if edit_title {
                        entry.title = newtext;
                    } else {
                        entry.description = newtext;
                    }
                    let _ = mddskmgr::config::save_atomic(&app.cfg, &app.cfg_paths);
                    eprintln!("quick_edit: saved config -> {:?}", app.cfg_paths.cfg_file);
                    update_overlay_text(app);
                    updated = true;
                }
            });
            if updated {
                refresh_visibility_now();
            }
        }
    }
}

extern "system" fn wndproc(hwnd: HWND, msg: u32, w: WPARAM, l: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            APP.with(|slot| {
                let (cfg, paths) = config::load_or_default().expect("config load");
                let overlay = Overlay::new(hwnd, &cfg.appearance.font_family, cfg.appearance.font_size_dip).expect("overlay");
                let taskbar_created_msg = unsafe { RegisterWindowMessageW(PCWSTR(windows::core::w!("TaskbarCreated").as_wide().as_ptr())) };
                let tray = Tray::new(hwnd, "Desktop Overlay").expect("tray");

                // Register hotkeys (warn on duplicates)
                let hk = &cfg.hotkeys;
                if mddskmgr::hotkeys::has_duplicates(hk) {
                    // Show a friendly tray balloon (without holding a RefCell borrow).
                    let _ = mddskmgr::tray::Tray::balloon_for(hwnd, "Hotkeys", "Duplicate hotkeys detected; adjust labels.json");
                }
                let _ = hotkeys::register(hwnd, hk.edit_title.ctrl, hk.edit_title.alt, hk.edit_title.shift, &hk.edit_title.key, HK_EDIT_TITLE);
                let _ = hotkeys::register(hwnd, hk.edit_description.ctrl, hk.edit_description.alt, hk.edit_description.shift, &hk.edit_description.key, HK_EDIT_DESC);
                let _ = hotkeys::register(hwnd, hk.toggle_overlay.ctrl, hk.toggle_overlay.alt, hk.toggle_overlay.shift, &hk.toggle_overlay.key, HK_TOGGLE);

                let current_guid = vd::get_current_desktop_guid();
                let vd_thread = mddskmgr::vd::start_vd_events(hwnd, WM_VD_SWITCHED);
                let mut app = AppState { hwnd, cfg, cfg_paths: paths, overlay, current_guid, visible: true, tray, taskbar_created_msg, vd_thread, hide_for_accessibility: false, hide_for_fullscreen: false };
                update_overlay_text(&mut app);
                *slot.borrow_mut() = Some(app);

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
        WM_VD_SWITCHED => {
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    let id = vd::get_current_desktop_guid();
                    if id != app.current_guid {
                        app.current_guid = id;
                        update_overlay_text(app);
                    }
                }
            });
            LRESULT(0)
        }
        WM_CFG_CHANGED => {
            // Reload config and apply labels/hotkeys; show any balloon outside borrow.
            let mut need_balloon = false;
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    if let Ok((new_cfg, _)) = mddskmgr::config::load_or_default() {
                        app.cfg = new_cfg;
                        // Re-register hotkeys
                        mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_TITLE);
                        mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_DESC);
                        mddskmgr::hotkeys::unregister(app.hwnd, HK_TOGGLE);
                        let hk = &app.cfg.hotkeys;
                        let ok1 = mddskmgr::hotkeys::register(app.hwnd, hk.edit_title.ctrl, hk.edit_title.alt, hk.edit_title.shift, &hk.edit_title.key, HK_EDIT_TITLE).unwrap_or(false);
                        let ok2 = mddskmgr::hotkeys::register(app.hwnd, hk.edit_description.ctrl, hk.edit_description.alt, hk.edit_description.shift, &hk.edit_description.key, HK_EDIT_DESC).unwrap_or(false);
                        let ok3 = mddskmgr::hotkeys::register(app.hwnd, hk.toggle_overlay.ctrl, hk.toggle_overlay.alt, hk.toggle_overlay.shift, &hk.toggle_overlay.key, HK_TOGGLE).unwrap_or(false);
                        if !(ok1 && ok2 && ok3) { need_balloon = true; }
                        update_overlay_text(app);
                    }
                }
            });
            if need_balloon {
                let _ = mddskmgr::tray::Tray::balloon_for(hwnd, "Hotkeys", "Some hotkeys failed to register. Adjust in labels.json");
            }
            LRESULT(0)
        }
        WM_TIMER => {
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    if w.0 == 1 { // VD poller
                        let id = vd::get_current_desktop_guid();
                        if id != app.current_guid {
                            app.current_guid = id;
                            update_overlay_text(app);
                        }
                    } else if w.0 == 2 { // visibility check
                        let hide = is_foreground_fullscreen(app);
                        app.hide_for_fullscreen = hide;
                    }
                }
            });
            if w.0 == 2 { refresh_visibility_now(); }
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
                        if let Some(app) = &*slot.borrow() {
                            Some(app.cfg_paths.cfg_file.to_string_lossy().to_string())
                        } else { None }
                    });
                    if let Some(path) = path {
                        let wpath: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
                        unsafe { let _ = ShellExecuteW(None, PCWSTR(windows::core::w!("open").as_wide().as_ptr()), PCWSTR(wpath.as_ptr()), None, None, SW_SHOWNORMAL); }
                    }
                }
                CMD_EXIT => unsafe { PostQuitMessage(0); },
                _ => {}
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            APP.with(|slot| {
                if let Some(app) = &mut *slot.borrow_mut() {
                    mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_TITLE);
                    mddskmgr::hotkeys::unregister(app.hwnd, HK_EDIT_DESC);
                    mddskmgr::hotkeys::unregister(app.hwnd, HK_TOGGLE);
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
    // File logging (daily) + env filter
    let (cfg_paths_log, _): (String, ()) = {
        let p = mddskmgr::config::project_paths().ok();
        if let Some(paths) = p {
            (paths.log_dir.to_string_lossy().to_string(), ())
        } else {
            (".".to_string(), ())
        }
    };
    std::fs::create_dir_all(&cfg_paths_log).ok();
    let file_appender = tracing_appender::rolling::daily(&cfg_paths_log, "overlay.log");
    let (nb, _guard) = tracing_appender::non_blocking(file_appender);
    tracing_subscriber::fmt()
        .with_env_filter(tracing_subscriber::EnvFilter::from_default_env())
        .with_writer(nb)
        .init();

    if !single_instance_guard() {
        println!("Another instance is already running. Exiting.");
        return Ok(());
    }

    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;

        let class_name = windows::core::w!("DesktopOverlayWndClass");
        let hinst = GetModuleHandleW(None).unwrap();
        let wc = WNDCLASSW {
            lpfnWndProc: Some(wndproc),
            hInstance: hinst.into(),
            lpszClassName: class_name,
            ..Default::default()
        };
        RegisterClassW(&wc);

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE((WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_TOPMOST | WS_EX_NOACTIVATE).0),
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
        // Pin overlay window across desktops when supported
        let _ = winvd::pin_window(hwnd);
        let _ = ShowWindow(hwnd, SW_SHOW);

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
