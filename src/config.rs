use anyhow::{Context, Result};
use directories::ProjectDirs;
use serde::{Deserialize, Serialize};
use std::{collections::HashMap, fs, io::Write, path::PathBuf};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    pub desktops: HashMap<String, DesktopLabel>,
    pub hotkeys: Hotkeys,
    pub appearance: Appearance,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct DesktopLabel {
    pub title: String,
    pub description: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Hotkeys {
    pub edit_title: KeyChord,
    pub edit_description: KeyChord,
    pub toggle_overlay: KeyChord,
    #[serde(default = "default_snap_key")]
    pub snap_position: KeyChord,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct KeyChord {
    pub ctrl: bool,
    pub alt: bool,
    pub shift: bool,
    pub key: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Appearance {
    pub font_family: String,
    pub font_size_dip: u32,
    pub margin_px: i32,
    #[serde(default)]
    pub hide_on_fullscreen: bool,
}

#[derive(Debug, Clone)]
pub struct Paths {
    pub cfg_file: PathBuf,
    pub cfg_dir: PathBuf,
    pub log_dir: PathBuf,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            desktops: HashMap::new(),
            hotkeys: Hotkeys {
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
            },
            appearance: Appearance {
                font_family: "Segoe UI".into(),
                font_size_dip: 16,
                margin_px: 8,
                hide_on_fullscreen: false,
            },
        }
    }
}

fn default_snap_key() -> KeyChord {
    KeyChord {
        ctrl: true,
        alt: true,
        shift: false,
        key: "L".into(),
    }
}

pub fn project_paths() -> Result<Paths> {
    let dirs = ProjectDirs::from("com", "Acme", "DesktopLabeler")
        .context("Failed to determine project directories")?;
    let cfg_dir = dirs.config_dir().to_path_buf();
    let cfg_file = cfg_dir.join("labels.json");
    let log_dir = dirs.data_local_dir().join("logs");
    Ok(Paths {
        cfg_file,
        cfg_dir,
        log_dir,
    })
}

pub fn load_or_default() -> Result<(Config, Paths)> {
    let paths = project_paths()?;
    fs::create_dir_all(&paths.cfg_dir).ok();
    fs::create_dir_all(&paths.log_dir).ok();
    let cfg = match fs::read_to_string(&paths.cfg_file) {
        Ok(s) => serde_json::from_str(&s).unwrap_or_default(),
        Err(_) => {
            // Migrate from old app name if present
            if let Some(old_dirs) = ProjectDirs::from("com", "Acme", "DesktopOverlay") {
                let old_file = old_dirs.config_dir().join("labels.json");
                if let Ok(s) = fs::read_to_string(&old_file) {
                    let parsed: Config = serde_json::from_str(&s).unwrap_or_default();
                    // Save to new location
                    let _ = save_atomic(&parsed, &paths);
                    parsed
                } else {
                    Config::default()
                }
            } else {
                Config::default()
            }
        }
    };
    Ok((cfg, paths))
}

pub fn save_atomic(cfg: &Config, paths: &Paths) -> Result<()> {
    fs::create_dir_all(&paths.cfg_dir).ok();
    let tmp = paths.cfg_file.with_extension("json.tmp");
    let data = serde_json::to_vec_pretty(cfg)?;
    {
        let mut f = fs::File::create(&tmp).context("create temp cfg")?;
        f.write_all(&data).context("write temp cfg")?;
        f.sync_all().ok();
    }
    // Best-effort atomic replace.
    fs::rename(&tmp, &paths.cfg_file).context("rename temp to final")?;
    Ok(())
}
