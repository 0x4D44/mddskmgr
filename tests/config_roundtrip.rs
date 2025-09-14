use mddskmgr::config::{Appearance, Config, DesktopLabel, Hotkeys, KeyChord, Paths, save_atomic};
use pretty_assertions::assert_eq;
use std::fs;

#[test]
fn save_and_load_roundtrip() {
    let mut cfg = Config::default();
    cfg.desktops.insert(
        "guid-1".into(),
        DesktopLabel {
            title: "Work".into(),
            description: "Tickets".into(),
        },
    );
    cfg.hotkeys = Hotkeys {
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
            key: "L".into(),
        },
    };
    cfg.appearance = Appearance {
        font_family: "Segoe UI".into(),
        font_size_dip: 16,
        margin_px: 8,
        hide_on_fullscreen: false,
    };

    let td = tempfile::tempdir().expect("tmpdir");
    let base = td.path();
    let cfg_dir = base.join("cfg");
    let log_dir = base.join("log");
    fs::create_dir_all(&cfg_dir).unwrap();
    fs::create_dir_all(&log_dir).unwrap();
    let paths = Paths {
        cfg_file: cfg_dir.join("labels.json"),
        cfg_dir,
        log_dir,
    };
    save_atomic(&cfg, &paths).expect("save");
    let data = fs::read_to_string(&paths.cfg_file).expect("read file");
    let parsed: Config = serde_json::from_str(&data).expect("json");
    assert_eq!(parsed.desktops.get("guid-1").unwrap().title, "Work");
    assert_eq!(parsed.hotkeys.toggle_overlay.key, "O");
}
