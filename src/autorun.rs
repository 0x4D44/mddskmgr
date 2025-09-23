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
fn startup_approved_key() -> anyhow::Result<HKEY> {
    let mut hkey = HKEY::default();
    let status = unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            windows::core::w!(
                "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartupApproved\\Run"
            ),
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
        Err(anyhow::anyhow!(
            "RegCreateKeyExW StartupApproved failed: {:?}",
            status
        ))
    }
}

#[cfg(windows)]
#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn registry_value_exists(
    hkey: HKEY,
    value: windows::core::PCWSTR,
    expected_type: REG_VALUE_TYPE,
) -> bool {
    let mut ty = REG_VALUE_TYPE(0);
    let mut cb = 0u32;
    unsafe { RegQueryValueExW(hkey, value, None, Some(&mut ty), None, Some(&mut cb)) }.is_ok()
        && ty == expected_type
        && cb > 0
}

#[cfg(windows)]
#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn read_binary_value(hkey: HKEY, value: windows::core::PCWSTR) -> Option<Vec<u8>> {
    let mut ty = REG_VALUE_TYPE(0);
    let mut cb = 0u32;
    if unsafe { RegQueryValueExW(hkey, value, None, Some(&mut ty), None, Some(&mut cb)) }.is_err()
        || ty != REG_BINARY
        || cb == 0
    {
        return None;
    }
    let mut buf = vec![0u8; cb as usize];
    let mut cb2 = cb;
    if unsafe {
        RegQueryValueExW(
            hkey,
            value,
            None,
            None,
            Some(buf.as_mut_ptr()),
            Some(&mut cb2),
        )
    }
    .is_err()
    {
        return None;
    }
    buf.truncate(cb2 as usize);
    Some(buf)
}

#[cfg(windows)]
#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn ensure_startup_marker(
    hkey: HKEY,
    value: windows::core::PCWSTR,
    enabled: bool,
) -> anyhow::Result<()> {
    let mut data = unsafe { read_binary_value(hkey, value) }.unwrap_or_else(|| vec![0u8; 8]);
    if data.is_empty() {
        data.resize(8, 0);
    }
    data[0] = if enabled { 0x02 } else { 0x03 };
    let status = unsafe { RegSetValueExW(hkey, value, 0, REG_BINARY, Some(&data)) };
    if status.is_err() {
        return Err(anyhow::anyhow!(
            "RegSetValueExW StartupApproved failed: {:?}",
            status
        ));
    }
    Ok(())
}

#[cfg(windows)]
#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn startup_entry_disabled(hkey: HKEY, value: windows::core::PCWSTR) -> bool {
    if let Some(data) = unsafe { read_binary_value(hkey, value) } {
        data.first().copied() == Some(0x03)
    } else {
        false
    }
}

#[cfg(windows)]
pub fn get_run_at_login() -> bool {
    unsafe {
        let mut has_run_value = false;
        if let Ok(hkey) = run_key() {
            has_run_value |=
                registry_value_exists(hkey, windows::core::w!("DesktopLabeler"), REG_SZ);
            has_run_value |=
                registry_value_exists(hkey, windows::core::w!("DesktopNameManager"), REG_SZ);
            let _ = RegCloseKey(hkey);
        }
        if !has_run_value {
            return false;
        }
        if let Ok(hkey) = startup_approved_key() {
            let disabled = startup_entry_disabled(hkey, windows::core::w!("DesktopLabeler"))
                || startup_entry_disabled(hkey, windows::core::w!("DesktopNameManager"));
            let _ = RegCloseKey(hkey);
            if disabled {
                return false;
            }
        }
        true
    }
}

#[cfg(windows)]
pub fn set_run_at_login(enable: bool) -> anyhow::Result<()> {
    unsafe {
        let h_run = run_key()?;
        let h_start = startup_approved_key()?;
        let result = (|| {
            if enable {
                let exe = std::env::current_exe()?;
                let val = format!("\"{}\"", exe.display());
                let data: Vec<u8> = to_utf16(&val)
                    .into_iter()
                    .flat_map(|u| u.to_le_bytes())
                    .collect();
                let status = RegSetValueExW(
                    h_run,
                    windows::core::w!("DesktopLabeler"),
                    0,
                    REG_SZ,
                    Some(&data),
                );
                if status.is_err() {
                    return Err(anyhow::anyhow!("RegSetValueExW failed: {:?}", status));
                }
                let _ = RegDeleteValueW(h_run, windows::core::w!("DesktopNameManager"));
                ensure_startup_marker(h_start, windows::core::w!("DesktopLabeler"), true)?;
                let _ = RegDeleteValueW(h_start, windows::core::w!("DesktopNameManager"));
            } else {
                let _ = RegDeleteValueW(h_run, windows::core::w!("DesktopLabeler"));
                let _ = RegDeleteValueW(h_run, windows::core::w!("DesktopNameManager"));
                ensure_startup_marker(h_start, windows::core::w!("DesktopLabeler"), false)?;
                let _ = RegDeleteValueW(h_start, windows::core::w!("DesktopNameManager"));
            }
            Ok(())
        })();
        let _ = RegCloseKey(h_start);
        let _ = RegCloseKey(h_run);
        result
    }
}

#[cfg(not(windows))]
pub fn get_run_at_login() -> bool {
    false
}
#[cfg(not(windows))]
pub fn set_run_at_login(_enable: bool) -> anyhow::Result<()> {
    Ok(())
}
