use anyhow::Result;
#[cfg(windows)]
use windows::Win32::Foundation::HWND;
#[cfg(windows)]
use windows::Win32::UI::Input::KeyboardAndMouse::{
    HOT_KEY_MODIFIERS, MOD_ALT, MOD_CONTROL, MOD_SHIFT, RegisterHotKey, UnregisterHotKey,
};

#[derive(Debug, Clone, Copy)]
pub struct Vk(pub u32);

pub fn vk_from_char(c: &str) -> Vk {
    // Accept single ASCII letter/digit; fall back to VK_BACK as harmless default.
    let up = c.trim().to_uppercase();
    let code = up.bytes().next().unwrap_or(b'B') as u32;
    Vk(code)
}

#[cfg(windows)]
pub fn register(
    hwnd: HWND,
    ctrl: bool,
    alt: bool,
    shift: bool,
    key: &str,
    id: i32,
) -> Result<bool> {
    let mut mods = HOT_KEY_MODIFIERS(0);
    if ctrl {
        mods |= MOD_CONTROL;
    }
    if alt {
        mods |= MOD_ALT;
    }
    if shift {
        mods |= MOD_SHIFT;
    }
    let Vk(vk) = vk_from_char(key);
    let res = unsafe { RegisterHotKey(hwnd, id, mods, vk) };
    Ok(res.is_ok())
}

#[cfg(not(windows))]
pub fn register(
    _hwnd: (),
    _ctrl: bool,
    _alt: bool,
    _shift: bool,
    _key: &str,
    _id: i32,
) -> Result<bool> {
    Ok(true)
}

#[cfg(windows)]
pub fn unregister(hwnd: HWND, id: i32) {
    unsafe {
        let _ = UnregisterHotKey(hwnd, id);
    }
}

#[cfg(not(windows))]
pub fn unregister(_hwnd: (), _id: i32) {}

pub const HK_EDIT_TITLE: i32 = 1;
pub const HK_EDIT_DESC: i32 = 2;
pub const HK_TOGGLE: i32 = 3;
pub const HK_SNAP: i32 = 4;

// Utility: detect duplicates between hotkey chords (case-insensitive key, same modifiers).
use crate::config::Hotkeys;
pub fn has_duplicates(hk: &Hotkeys) -> bool {
    fn same(a: &crate::config::KeyChord, b: &crate::config::KeyChord) -> bool {
        a.ctrl == b.ctrl
            && a.alt == b.alt
            && a.shift == b.shift
            && a.key.eq_ignore_ascii_case(&b.key)
    }
    same(&hk.edit_title, &hk.edit_description)
        || same(&hk.edit_title, &hk.toggle_overlay)
        || same(&hk.edit_description, &hk.toggle_overlay)
        || same(&hk.edit_title, &hk.snap_position)
        || same(&hk.edit_description, &hk.snap_position)
        || same(&hk.toggle_overlay, &hk.snap_position)
}
