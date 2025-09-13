use anyhow::Result;
use core::ffi::c_void;
use std::sync::mpsc;
use std::thread;
use std::time::Duration;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::PostMessageW;
use winvd::{DesktopEvent, DesktopEventThread, listen_desktop_events};

pub fn get_current_desktop_guid() -> String {
    #[cfg(windows)]
    {
        if let Ok(d) = winvd::get_current_desktop() {
            return format!("{:?}", d);
        }
    }
    // Fallback if API unavailable
    "default".to_string()
}

// Placeholder for event listener; an implementation can spawn a thread and
// post a custom window message to the UI thread when the GUID changes.
pub fn start_vd_listener() -> Result<()> {
    // TODO: integrate winvd::listen_desktop_events and PostMessage to UI.
    Ok(())
}

pub fn start_vd_poller(hwnd: HWND, msg: u32) {
    let hwnd_raw = hwnd.0 as usize; // make Send
    thread::spawn(move || {
        let mut last = super::vd::get_current_desktop_guid();
        loop {
            let now = super::vd::get_current_desktop_guid();
            if now != last {
                let target = HWND(hwnd_raw as *mut c_void);
                unsafe {
                    let _ = PostMessageW(target, msg, WPARAM(0), LPARAM(0));
                }
                last = now;
            }
            thread::sleep(Duration::from_millis(250));
        }
    });
}

pub fn start_vd_events(hwnd: HWND, msg: u32) -> Option<DesktopEventThread> {
    let (tx, rx) = mpsc::channel::<DesktopEvent>();
    let thread = match listen_desktop_events::<DesktopEvent, _>(tx) {
        Ok(t) => t,
        Err(_) => return None,
    };
    let hwnd_raw = hwnd.0 as usize;
    thread::spawn(move || {
        for evt in rx {
            match evt {
                DesktopEvent::DesktopChanged { .. } | DesktopEvent::WindowChanged(_) => unsafe {
                    let _ = PostMessageW(HWND(hwnd_raw as *mut c_void), msg, WPARAM(0), LPARAM(0));
                },
                _ => {}
            }
        }
    });
    Some(thread)
}
