#[cfg(windows)]
use windows::Win32::System::Registry::*;

#[cfg(windows)]
fn to_utf16(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(windows)]
fn run_key() -> anyhow::Result<HKEY> {
    let mut hkey = HKEY::default();
    let status = unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            windows::core::w!("Software\\Microsoft\\Windows\\CurrentVersion\\Run"),
            0,
            None,
            REG_OPTION_NON_VOLATILE,
            KEY_READ | KEY_WRITE,
            None,
            &mut hkey,
            None,
        )
    };
    if status.is_ok() {
        Ok(hkey)
    } else {
        Err(anyhow::anyhow!("RegCreateKeyExW failed: {:?}", status))
    }
}

#[cfg(windows)]
pub fn get_run_at_login() -> bool {
    unsafe {
        if let Ok(hkey) = run_key() {
            let mut ty = REG_VALUE_TYPE(0);
            let mut cb: u32 = 0;
            if RegQueryValueExW(
                hkey,
                windows::core::w!("DesktopLabeler"),
                None,
                Some(&mut ty),
                None,
                Some(&mut cb),
            )
            .is_ok()
                && ty == REG_SZ
                && cb > 0
            {
                return true;
            }
            // Also accept legacy value name (migrate path but keep behavior)
            let mut ty2 = REG_VALUE_TYPE(0);
            let mut cb2: u32 = 0;
            if RegQueryValueExW(
                hkey,
                windows::core::w!("DesktopNameManager"),
                None,
                Some(&mut ty2),
                None,
                Some(&mut cb2),
            )
            .is_ok()
                && ty2 == REG_SZ
                && cb2 > 0
            {
                return true;
            }
        }
        false
    }
}

#[cfg(windows)]
pub fn set_run_at_login(enable: bool) -> anyhow::Result<()> {
    unsafe {
        let hkey = run_key()?;
        if enable {
            let exe = std::env::current_exe()?;
            let val = format!("\"{}\"", exe.display());
            let data: Vec<u8> = to_utf16(&val)
                .into_iter()
                .flat_map(|u| u.to_le_bytes())
                .collect();
            let status = RegSetValueExW(
                hkey,
                windows::core::w!("DesktopLabeler"),
                0,
                REG_SZ,
                Some(&data),
            );
            if status.is_err() {
                return Err(anyhow::anyhow!("RegSetValueExW failed: {:?}", status));
            }
            // Remove legacy entry if present
            let _ = RegDeleteValueW(hkey, windows::core::w!("DesktopNameManager"));
        } else {
            let _ = RegDeleteValueW(hkey, windows::core::w!("DesktopLabeler"));
        }
    }
    Ok(())
}

#[cfg(not(windows))]
pub fn get_run_at_login() -> bool {
    false
}
#[cfg(not(windows))]
pub fn set_run_at_login(_enable: bool) -> anyhow::Result<()> {
    Ok(())
}
