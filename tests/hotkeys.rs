use mddskmgr::hotkeys::vk_from_char;

#[test]
fn maps_alpha_keys_to_vk() {
    let t = vk_from_char("t");
    assert_eq!(t.0, 'T' as u32);
    let d = vk_from_char("D");
    assert_eq!(d.0, 'D' as u32);
}

#[test]
fn default_key_when_empty() {
    let b = vk_from_char("");
    assert_eq!(b.0, 'B' as u32);
}
