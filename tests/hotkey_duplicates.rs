use mddskmgr::config::{Hotkeys, KeyChord};

fn dup(a: &KeyChord, b: &KeyChord) -> bool {
    a.ctrl == b.ctrl && a.alt == b.alt && a.shift == b.shift && a.key.eq_ignore_ascii_case(&b.key)
}

fn has_duplicates(hk: &Hotkeys) -> bool {
    dup(&hk.edit_title, &hk.edit_description)
        || dup(&hk.edit_title, &hk.toggle_overlay)
        || dup(&hk.edit_description, &hk.toggle_overlay)
        || dup(&hk.edit_title, &hk.snap_position)
        || dup(&hk.edit_description, &hk.snap_position)
        || dup(&hk.toggle_overlay, &hk.snap_position)
}

#[test]
fn detects_duplicate_hotkeys() {
    let mut hk = Hotkeys {
        edit_title: KeyChord {
            ctrl: true,
            alt: true,
            shift: false,
            key: "T".into(),
        },
        edit_description: KeyChord {
            ctrl: true,
            alt: true,
            shift: false,
            key: "D".into(),
        },
        toggle_overlay: KeyChord {
            ctrl: true,
            alt: true,
            shift: false,
            key: "O".into(),
        },
        snap_position: KeyChord {
            ctrl: true,
            alt: true,
            shift: false,
            key: "S".into(),
        },
    };
    assert!(!has_duplicates(&hk));
    // Collide description with title
    hk.edit_description.key = "t".into();
    assert!(has_duplicates(&hk));
}
