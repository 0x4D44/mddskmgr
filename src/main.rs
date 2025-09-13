// Windows-only implementation lives in src/windows_main.rs
#[cfg(windows)]
mod windows_main;

// Windows entry point calls into module
#[cfg(windows)]
fn main() -> anyhow::Result<()> {
    windows_main::main()
}

// Non-Windows stub builds cleanly and informs the user.
#[cfg(not(windows))]
fn main() {
    println!("mddskmgr is Windows-only. Build on Windows to run.");
}
