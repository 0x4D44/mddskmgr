pub mod autorun;
pub mod config;
pub mod core;
pub mod hotkeys;

// Windows-only modules
#[cfg(windows)]
pub mod overlay;
#[cfg(windows)]
pub mod tray;
#[cfg(windows)]
pub mod ui;
#[cfg(windows)]
pub mod vd;
