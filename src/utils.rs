/// Convert a string to a null-terminated UTF-16 vector suitable for Windows API calls.
pub fn to_utf16(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}
