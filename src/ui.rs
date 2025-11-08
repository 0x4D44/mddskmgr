use crate::utils::to_utf16;
use core::ffi::c_void;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Foundation::{LPARAM as LPARAM_T, WPARAM as WPARAM_T};
use windows::Win32::Graphics::Gdi::{DEFAULT_GUI_FONT, GetStockObject};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::HiDpi::{AdjustWindowRectExForDpi, GetDpiForWindow};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::PCWSTR;

// Fallback FFI for SetFocus: windows crate may not expose it in all builds
#[cfg(windows)]
#[link(name = "user32")]
unsafe extern "system" {
    fn SetFocus(hWnd: HWND) -> HWND;
}

struct DialogState {
    text: String,
    hint: String,
    done: bool,
    edit_hwnd: HWND,
}

const EM_LIMITTEXT: u32 = 0x00C5;
const EM_SETSEL: u32 = 0x00B1;
const SS_LEFT: u32 = 0x0000;

fn scale(dpi: u32, v: i32) -> i32 {
    ((v as i64 * dpi as i64 + 48) / 96) as i32
}

#[allow(unknown_lints)]
#[allow(clippy::manual_dangling_ptr)]
fn menu_id(id: usize) -> HMENU {
    HMENU(id as *mut c_void)
}

#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn set_ctrl_font(hwnd: HWND) {
    let hobj = unsafe { GetStockObject(DEFAULT_GUI_FONT) };
    let _ = unsafe { SendMessageW(hwnd, WM_SETFONT, WPARAM_T(hobj.0 as usize), LPARAM_T(1)) };
}

#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn layout_dialog(hwnd: HWND) {
    let dpi = GetDpiForWindow(hwnd);
    let margin = scale(dpi, 12);
    let gap = scale(dpi, 8);
    let label_h = scale(dpi, 20);
    let edit_h = scale(dpi, 28);
    let btn_w = scale(dpi, 88);
    let btn_h = scale(dpi, 28);
    let client_w = scale(dpi, 460);

    // Controls
    let hlabel = GetDlgItem(hwnd, 1000)
        .ok()
        .unwrap_or(HWND(std::ptr::null_mut()));
    let hedit = GetDlgItem(hwnd, 1001)
        .ok()
        .unwrap_or(HWND(std::ptr::null_mut()));
    let hok = GetDlgItem(hwnd, 1)
        .ok()
        .unwrap_or(HWND(std::ptr::null_mut()));
    let hcancel = GetDlgItem(hwnd, 2)
        .ok()
        .unwrap_or(HWND(std::ptr::null_mut()));

    let y_label = margin;
    let y_edit = y_label + label_h + gap;
    let y_btn = y_edit + edit_h + gap + scale(dpi, 6);

    let ok_x = client_w - margin - btn_w * 2 - gap;
    let cancel_x = client_w - margin - btn_w;

    if !hlabel.0.is_null() {
        let _ = MoveWindow(
            hlabel,
            margin,
            y_label,
            client_w - margin * 2,
            label_h,
            true,
        );
    }
    if !hedit.0.is_null() {
        let _ = MoveWindow(hedit, margin, y_edit, client_w - margin * 2, edit_h, true);
    }
    if !hok.0.is_null() {
        let _ = MoveWindow(hok, ok_x, y_btn, btn_w, btn_h, true);
    }
    if !hcancel.0.is_null() {
        let _ = MoveWindow(hcancel, cancel_x, y_btn, btn_w, btn_h, true);
    }

    // Set fonts
    if !hlabel.0.is_null() {
        set_ctrl_font(hlabel);
    }
    if !hedit.0.is_null() {
        set_ctrl_font(hedit);
    }
    if !hok.0.is_null() {
        set_ctrl_font(hok);
    }
    if !hcancel.0.is_null() {
        set_ctrl_font(hcancel);
    }

    // Resize window to fit client + margins
    let client_h = y_btn + btn_h + margin;
    let mut rc = RECT {
        left: 0,
        top: 0,
        right: client_w,
        bottom: client_h,
    };
    let _ = AdjustWindowRectExForDpi(
        &mut rc,
        WS_CAPTION | WS_SYSMENU | WS_POPUPWINDOW,
        false,
        WINDOW_EX_STYLE(WS_EX_TOOLWINDOW.0 | WS_EX_TOPMOST.0),
        dpi,
    );
    let wnd_w = rc.right - rc.left;
    let wnd_h = rc.bottom - rc.top;
    let _ = SetWindowPos(
        hwnd,
        None,
        0,
        0,
        wnd_w,
        wnd_h,
        SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE,
    );
}

pub fn prompt_text(parent: HWND, caption: &str, hint: &str, initial: &str) -> Option<String> {
    unsafe {
        tracing::debug!(caption=%caption, hint=%hint, initial=%initial, "prompt_text");
        let class = windows::core::w!("OverlayInputDlg");
        let hinst = GetModuleHandleW(None).unwrap();
        let wc = WNDCLASSW {
            lpfnWndProc: Some(dlg_wndproc),
            hInstance: hinst.into(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            lpszClassName: class,
            ..Default::default()
        };
        // Ignore error if already registered
        let _ = RegisterClassW(&wc);

        // Center 420x140 window near the parent
        let (w, h) = (420, 140);
        let (x, y) = center_on_parent(parent, w, h);
        // Remember previous foreground to restore later
        let prev_fg = GetForegroundWindow();
        // Prepare initial state and pass pointer via lpParam so WM_CREATE can use it
        let state = Box::new(DialogState {
            text: initial.to_string(),
            hint: hint.to_string(),
            done: false,
            edit_hwnd: HWND(0 as _),
        });
        let state_ptr = Box::into_raw(state);

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE(WS_EX_TOOLWINDOW.0 | WS_EX_TOPMOST.0 | WS_EX_CONTROLPARENT.0),
            class,
            PCWSTR(to_utf16(caption).as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_POPUPWINDOW,
            x,
            y,
            w,
            h,
            parent,
            None,
            hinst,
            Some(state_ptr as *mut core::ffi::c_void),
        )
        .ok()?;

        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = SetForegroundWindow(hwnd);

        // Modal loop
        let mut msg = MSG::default();
        loop {
            if GetMessageW(&mut msg, HWND(0 as _), 0, 0).into() {
                // Give dialog manager a chance to process VK_RETURN/VK_ESCAPE, tabbing, etc.
                if IsDialogMessageW(hwnd, &msg).as_bool() {
                    // fallthrough to check done
                } else {
                    let _ = TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            } else {
                break;
            }
            // Check state via preserved pointer (works even after DestroyWindow)
            if !state_ptr.is_null() && (*state_ptr).done {
                let boxed = Box::from_raw(state_ptr);
                let res = if boxed.text.is_empty() {
                    None
                } else {
                    Some(boxed.text)
                };
                tracing::debug!(res=?res.as_deref(), "prompt_text: returning");
                // Restore previous foreground window if valid
                if !prev_fg.0.is_null() && prev_fg != hwnd {
                    let _ = SetForegroundWindow(prev_fg);
                }
                return res;
            }
        }
        // If we broke out (rare), clean up and return None
        if !state_ptr.is_null() {
            let _ = Box::from_raw(state_ptr);
        }
        if !prev_fg.0.is_null() && prev_fg != hwnd {
            let _ = SetForegroundWindow(prev_fg);
        }
        None
    }
}

extern "system" fn dlg_wndproc(hwnd: HWND, msg: u32, w: WPARAM, l: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                // Stash DialogState from CREATESTRUCTW
                let cs: &CREATESTRUCTW = &*(l.0 as *const CREATESTRUCTW);
                if !cs.lpCreateParams.is_null() {
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, cs.lpCreateParams as isize);
                }
                tracing::debug!("dlg: WM_CREATE");
                // Create edit and buttons
                let hinst = GetModuleHandleW(None).unwrap();
                let style = WINDOW_STYLE(
                    WS_CHILD.0
                        | WS_VISIBLE.0
                        | WS_BORDER.0
                        | WS_TABSTOP.0
                        | (ES_LEFT as u32)
                        | (ES_AUTOHSCROLL as u32),
                );
                // Static hint label
                let _label = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(windows::core::w!("STATIC").as_wide().as_ptr()),
                    PCWSTR(to_utf16("").as_ptr()),
                    WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | SS_LEFT),
                    0,0,0,0,
                    hwnd,
                    menu_id(1000),
                    hinst,
                    None,
                );
                let edit = CreateWindowExW(
                    WINDOW_EX_STYLE(WS_EX_CLIENTEDGE.0),
                    PCWSTR(windows::core::w!("EDIT").as_wide().as_ptr()),
                    PCWSTR(to_utf16("").as_ptr()),
                    style,
                    0,0,0,0,
                    hwnd,
                    menu_id(1001),
                    hinst,
                    None,
                )
                .unwrap();
                // Limit text length to 200 chars
                let _ = SendMessageW(edit, EM_LIMITTEXT, WPARAM_T(200), LPARAM_T(0));
                let _ok = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(windows::core::w!("BUTTON").as_wide().as_ptr()),
                    PCWSTR(windows::core::w!("OK").as_wide().as_ptr()),
                    WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0 | (BS_PUSHBUTTON as u32)),
                    0,0,0,0,
                    hwnd,
                    menu_id(1),
                    hinst,
                    None,
                )
                .unwrap();
                let _cancel = CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    PCWSTR(windows::core::w!("BUTTON").as_wide().as_ptr()),
                    PCWSTR(windows::core::w!("Cancel").as_wide().as_ptr()),
                    WINDOW_STYLE(WS_CHILD.0 | WS_VISIBLE.0 | WS_TABSTOP.0),
                    0,0,0,0,
                    hwnd,
                    menu_id(2),
                    hinst,
                    None,
                )
                .unwrap();
                // Initialize from state (hint + initial text) and focus edit
                let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DialogState;
                if !p.is_null() {
                    (*p).edit_hwnd = edit;
                    let _ = SetWindowTextW(edit, PCWSTR(to_utf16(&(*p).text).as_ptr()));
                    layout_dialog(hwnd);
                    // Explicitly focus the edit control
                    let _ = SetFocus(edit);
                    // Select all text to ease replacement
                    let _ = SendMessageW(edit, EM_SETSEL, WPARAM_T(0), LPARAM_T(-1));
                    // Set label text
                    if let Ok(label_hwnd) = GetDlgItem(hwnd, 1000) {
                        let _ = SetWindowTextW(label_hwnd, PCWSTR(to_utf16(&(*p).hint).as_ptr()));
                    }
                }
                LRESULT(0)
            }
            0x02E0 /* WM_DPICHANGED */ => {
                layout_dialog(hwnd);
                LRESULT(0)
            }
            WM_SETFOCUS => {
                // Ensure the edit control gets focus if the dialog gains focus
                let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DialogState;
                if !p.is_null() && !(*p).edit_hwnd.0.is_null() {
                    let _ = SetFocus((*p).edit_hwnd);
                    let _ = SendMessageW((*p).edit_hwnd, EM_SETSEL, WPARAM_T(0), LPARAM_T(-1));
                    return LRESULT(0);
                }
                DefWindowProcW(hwnd, msg, w, l)
            }
            WM_COMMAND => {
                let id = (w.0 & 0xFFFF) as u16;
                match id {
                    1 => {
                        // Read text from edit control
                        let edit_hwnd: HWND = GetDlgItem(hwnd, 1001).unwrap();
                        let len = GetWindowTextLengthW(edit_hwnd);
                        let mut buf: Vec<u16> = vec![0; (len + 1) as usize];
                        let _ = GetWindowTextW(edit_hwnd, &mut buf);
                        let s = String::from_utf16_lossy(
                            &buf[..buf.iter().position(|&c| c == 0).unwrap_or(buf.len())],
                        );
                        tracing::debug!(text=%s, "dlg: OK pressed");
                        let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DialogState;
                        if !p.is_null() {
                            (*p).text = s;
                            (*p).done = true;
                        }
                        let _ = DestroyWindow(hwnd);
                        LRESULT(0)
                    }
                    2 => {
                        tracing::debug!("dlg: Cancel pressed");
                        let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DialogState;
                        if !p.is_null() {
                            (*p).text.clear();
                            (*p).done = true;
                        }
                        let _ = DestroyWindow(hwnd);
                        LRESULT(0)
                    }
                    _ => DefWindowProcW(hwnd, msg, w, l),
                }
            }
            // Fallback: handle Enter/Esc if IsDialogMessage didn't
            WM_KEYDOWN => {
                let vk = w.0 as u32;
                match vk {
                    0x0D => {
                        // VK_RETURN
                        let _ = SendMessageW(hwnd, WM_COMMAND, WPARAM_T(1), LPARAM_T(0));
                        LRESULT(0)
                    }
                    0x1B => {
                        // VK_ESCAPE
                        let _ = SendMessageW(hwnd, WM_COMMAND, WPARAM_T(2), LPARAM_T(0));
                        LRESULT(0)
                    }
                    _ => DefWindowProcW(hwnd, msg, w, l),
                }
            }
            WM_CLOSE => {
                let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DialogState;
                if !p.is_null() {
                    (*p).text.clear();
                    (*p).done = true;
                }
                let _ = DestroyWindow(hwnd);
                LRESULT(0)
            }
            WM_NCDESTROY => {
                // Do not free state here; prompt_text will reclaim and free it.
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, w, l),
        }
    }
}

fn center_on_parent(parent: HWND, w: i32, h: i32) -> (i32, i32) {
    unsafe {
        let mut rc: RECT = RECT::default();
        if parent.0.is_null() || GetWindowRect(parent, &mut rc).is_err() {
            let x = (GetSystemMetrics(SM_CXSCREEN) - w) / 2;
            let y = (GetSystemMetrics(SM_CYSCREEN) - h) / 2;
            return (x, y);
        }
        let cx = rc.left + ((rc.right - rc.left - w) / 2);
        let cy = rc.top + ((rc.bottom - rc.top - h) / 2);
        (cx, cy)
    }
}
