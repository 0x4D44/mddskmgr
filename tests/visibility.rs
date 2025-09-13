use mddskmgr::core::should_show;

#[test]
fn visibility_truth_table() {
    // toggled, hc, fs -> show?
    assert!(should_show(true, false, false));
    assert!(!should_show(true, true, false));
    assert!(!should_show(true, false, true));
    assert!(!should_show(true, true, true));
    assert!(!should_show(false, false, false));
    assert!(!should_show(false, true, false));
    assert!(!should_show(false, false, true));
}
