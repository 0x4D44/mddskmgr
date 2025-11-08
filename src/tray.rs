use crate::utils::to_utf16;
use anyhow::{Result, anyhow};
use std::mem::size_of;
use windows::Win32::Foundation::{HWND, POINT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::ExtractIconW;
use windows::Win32::UI::Shell::{
    NIF_ICON, NIF_INFO, NIF_MESSAGE, NIF_TIP, NIM_ADD, NIM_DELETE, NIM_MODIFY, NOTIFYICONDATAW,
    Shell_NotifyIconW,
};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::PCWSTR;

pub const TRAY_UID: u32 = 1;
pub const TRAY_MSG: u32 = WM_APP + 1;

pub const CMD_EDIT_TITLE: u16 = 1001;
pub const CMD_EDIT_DESC: u16 = 1002;
pub const CMD_TOGGLE: u16 = 1003;
pub const CMD_OPEN_CONFIG: u16 = 1004;
pub const CMD_EXIT: u16 = 1005;
pub const CMD_ABOUT: u16 = 1006;
pub const CMD_RUN_AT_STARTUP: u16 = 1007;

pub struct Tray {
    pub nid: NOTIFYICONDATAW,
}

impl Tray {
    fn load_app_icon() -> HICON {
        unsafe {
            let hinst = GetModuleHandleW(None).unwrap_or_default();
            // Try extracting the primary icon from our executable
            if let Ok(exe) = std::env::current_exe() {
                let wpath = to_utf16(&exe.display().to_string());
                let icon = ExtractIconW(hinst, PCWSTR(wpath.as_ptr()), 0);
                if !icon.0.is_null() && icon.0 as usize > 1 {
                    return icon;
                }
            }
            // Fallback to stock app icon
            LoadIconW(None, IDI_APPLICATION).unwrap_or_default()
        }
    }

    pub fn new(hwnd: HWND, tip: &str) -> Result<Self> {
        unsafe {
            // SAFETY: NOTIFYICONDATAW is a C struct from the Windows API that can be safely zero-initialized.
            // All fields are explicitly set below, and Windows API documentation specifies this initialization pattern.
            let mut nid: NOTIFYICONDATAW = std::mem::zeroed();
            nid.cbSize = size_of::<NOTIFYICONDATAW>() as u32;
            nid.hWnd = hwnd;
            nid.uID = TRAY_UID;
            nid.uFlags = NIF_MESSAGE | NIF_TIP | NIF_ICON;
            nid.uCallbackMessage = TRAY_MSG;
            // Load our embedded app icon; fallback to stock if needed
            nid.hIcon = Self::load_app_icon();
            // Set tooltip
            let wtip = to_utf16(tip);
            let lt = wtip.len().min(nid.szTip.len());
            nid.szTip[..lt].copy_from_slice(&wtip[..lt]);
            if !Shell_NotifyIconW(NIM_ADD, &nid).as_bool() {
                return Err(anyhow!("Shell_NotifyIconW(NIM_ADD) failed"));
            }
            Ok(Self { nid })
        }
    }

    pub fn remove_icon(&mut self) {
        unsafe {
            let _ = Shell_NotifyIconW(NIM_DELETE, &self.nid);
        }
    }

    pub fn re_add(&mut self) {
        unsafe {
            let _ = Shell_NotifyIconW(NIM_ADD, &self.nid);
        }
    }

    pub fn show_balloon(&mut self, title: &str, text: &str) {
        unsafe {
            self.nid.uFlags = NIF_INFO | NIF_TIP | NIF_MESSAGE | NIF_ICON;
            let wtitle = to_utf16(title);
            let wtext = to_utf16(text);
            let lt = wtitle.len().min(self.nid.szInfoTitle.len());
            self.nid.szInfoTitle[..lt].copy_from_slice(&wtitle[..lt]);
            let li = wtext.len().min(self.nid.szInfo.len());
            self.nid.szInfo[..li].copy_from_slice(&wtext[..li]);
            self.nid.hIcon = Self::load_app_icon();
            let _ = Shell_NotifyIconW(NIM_MODIFY, &self.nid);
        }
    }

    pub fn show_menu(&self, hwnd: HWND) -> Result<()> {
        unsafe {
            let hmenu = CreatePopupMenu()?;
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_EDIT_TITLE as usize,
                PCWSTR(windows::core::w!("Edit Title").as_wide().as_ptr()),
            )?;
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_EDIT_DESC as usize,
                PCWSTR(windows::core::w!("Edit Description").as_wide().as_ptr()),
            )?;
            AppendMenuW(hmenu, MF_SEPARATOR, 0, PCWSTR::null())?;
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_TOGGLE as usize,
                PCWSTR(windows::core::w!("Toggle Overlay").as_wide().as_ptr()),
            )?;
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_OPEN_CONFIG as usize,
                PCWSTR(windows::core::w!("Open Config").as_wide().as_ptr()),
            )?;
            AppendMenuW(hmenu, MF_SEPARATOR, 0, PCWSTR::null())?;
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_RUN_AT_STARTUP as usize,
                PCWSTR(windows::core::w!("Run at startup").as_wide().as_ptr()),
            )?;
            // Reflect current autorun state
            let enabled = crate::autorun::get_run_at_login();
            let _ = CheckMenuItem(
                hmenu,
                CMD_RUN_AT_STARTUP as u32,
                (MF_BYCOMMAND | if enabled { MF_CHECKED } else { MF_UNCHECKED }).0,
            );
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_ABOUT as usize,
                PCWSTR(windows::core::w!("About...").as_wide().as_ptr()),
            )?;
            AppendMenuW(
                hmenu,
                MF_STRING,
                CMD_EXIT as usize,
                PCWSTR(windows::core::w!("Exit").as_wide().as_ptr()),
            )?;

            let mut pt = POINT::default();
            let _ = GetCursorPos(&mut pt);
            let _ = SetForegroundWindow(hwnd);
            let _ = TrackPopupMenu(hmenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, None);
            let _ = DestroyMenu(hmenu);
        }
        Ok(())
    }

    // Static helpers to avoid borrowing AppState across re-entrant shell calls
    pub fn show_popup_menu(hwnd: HWND) -> Result<()> {
        Self {
            nid: unsafe { std::mem::zeroed() },
        }
        .show_menu(hwnd)
    }

    pub fn balloon_for(hwnd: HWND, title: &str, text: &str) -> Result<()> {
        unsafe {
            let mut nid: NOTIFYICONDATAW = std::mem::zeroed();
            nid.cbSize = size_of::<NOTIFYICONDATAW>() as u32;
            nid.hWnd = hwnd;
            nid.uID = TRAY_UID;
            nid.uFlags = NIF_INFO | NIF_ICON;
            let wtitle = to_utf16(title);
            let wtext = to_utf16(text);
            let lt = wtitle.len().min(nid.szInfoTitle.len());
            nid.szInfoTitle[..lt].copy_from_slice(&wtitle[..lt]);
            let li = wtext.len().min(nid.szInfo.len());
            nid.szInfo[..li].copy_from_slice(&wtext[..li]);
            nid.hIcon = Self::load_app_icon();
            if !Shell_NotifyIconW(NIM_MODIFY, &nid).as_bool() {
                // In case the icon is missing (e.g., Explorer restart), re-add then modify.
                nid.uFlags = NIF_MESSAGE | NIF_TIP | NIF_ICON | NIF_INFO;
                nid.uCallbackMessage = TRAY_MSG;
                nid.hIcon = Self::load_app_icon();
                let tip = "Desktop Labeler";
                let wtip = to_utf16(tip);
                let lt2 = wtip.len().min(nid.szTip.len());
                nid.szTip[..lt2].copy_from_slice(&wtip[..lt2]);
                let _ = Shell_NotifyIconW(NIM_ADD, &nid);
            }
        }
        Ok(())
    }

    pub fn re_add_for(hwnd: HWND) -> Result<()> {
        unsafe {
            let mut nid: NOTIFYICONDATAW = std::mem::zeroed();
            nid.cbSize = size_of::<NOTIFYICONDATAW>() as u32;
            nid.hWnd = hwnd;
            nid.uID = TRAY_UID;
            nid.uFlags = NIF_MESSAGE | NIF_TIP | NIF_ICON;
            nid.uCallbackMessage = TRAY_MSG;
            nid.hIcon = Self::load_app_icon();
            let tip = "Desktop Labeler";
            let wtip = to_utf16(tip);
            let lt = wtip.len().min(nid.szTip.len());
            nid.szTip[..lt].copy_from_slice(&wtip[..lt]);
            let _ = Shell_NotifyIconW(NIM_ADD, &nid);
        }
        Ok(())
    }
}
