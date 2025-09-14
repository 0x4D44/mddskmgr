# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Rust project called `mddskmgr` (binary: `mddsklbl`, version 1.0.0) using Rust edition 2024.

## Common Development Commands

### Build and Run
```bash
# Build the project
cargo build

# Build with optimizations
cargo build --release

# Run the project
cargo run --bin mddsklbl

# Run with release optimizations
cargo run --release --bin mddsklbl
```

### Code Quality
```bash
# Format code
cargo fmt

# Run clippy linter
cargo clippy

# Run clippy with all targets and strict warnings
cargo clippy --all-targets -- -D warnings
```

### Testing
```bash
# Run all tests
cargo test

# Run tests with output
cargo test -- --nocapture

# Run a specific test
cargo test test_name

# Run tests in release mode
cargo test --release
```

### Documentation
```bash
# Build and open documentation
cargo doc --open

# Build documentation without dependencies
cargo doc --no-deps
```

## Development Requirements

- Rust 1.85.0 or later
- All code must compile without warnings
- Code must pass `cargo fmt` and `cargo clippy` checks
