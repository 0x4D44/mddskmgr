#![cfg_attr(windows, windows_subsystem = "windows")]
// Windows-only implementation lives in src/windows_main.rs
#[cfg(windows)]
mod windows_main;

// Windows entry point: initialize logging then delegate to module
#[cfg(windows)]
fn main() -> anyhow::Result<()> {
    // Best-effort tracing setup to a rolling daily log under the app's data dir.
    // Falls back silently if initialization fails (e.g., IO errors).
    {
        if let Ok((_, paths)) = mddskmgr::config::load_or_default() {
            std::fs::create_dir_all(&paths.log_dir).ok();
            let file_appender = tracing_appender::rolling::daily(&paths.log_dir, "mddsklbl.log");
            let (nb_writer, _guard) = tracing_appender::non_blocking(file_appender);
            let env = tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| tracing_subscriber::EnvFilter::new("info"));
            let _ = tracing_subscriber::fmt()
                .with_env_filter(env)
                .with_ansi(false)
                .with_writer(nb_writer)
                .try_init();
            tracing::info!("mddsklbl starting");
        }
    }
    windows_main::main()
}

// Non-Windows stub builds cleanly and informs the user.
#[cfg(not(windows))]
fn main() {
    println!("Desktop Labeler (mddsklbl) is Windows-only. Build on Windows to run.");
}
