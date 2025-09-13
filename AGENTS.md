# Repository Guidelines

## Project Structure & Module Organization
- `Cargo.toml` — package metadata and dependencies.
- `src/main.rs` — binary entrypoint. Keep it thin; move logic into modules.
- Optional: `src/lib.rs` for reusable logic imported by `main.rs`.
- Optional: `tests/` for integration tests (one file per feature, e.g., `tests/cli_tests.rs`).

## Build, Test, and Development Commands
- `cargo build` — compile (debug).
- `cargo run -- <args>` — run the binary locally.
- `cargo test` — run unit + integration tests.
- `cargo check` — fast type/lint pass without building artifacts.
- `cargo fmt --all` and `cargo clippy -- -D warnings` — format and lint; required before PRs.

## Coding Style & Naming Conventions
- Use `rustfmt` defaults (4-space indent, standard wrapping). Do not hand‑format.
- Naming: modules/files and functions `snake_case`; types/traits `UpperCamelCase`; constants `SCREAMING_SNAKE_CASE`.
- Error handling: return `Result<_, _>` and propagate with `?`; keep `main.rs` focused on wiring (CLI/config), business logic in library modules.

## Testing Guidelines
- Unit tests live next to code under `#[cfg(test)] mod tests { ... }`.
- Integration tests go in `tests/`; name files by feature and tests descriptively: `does_<thing>_when_<condition>()`.
- Useful invocations: `cargo test -- --nocapture`, `cargo test module::case`.

## Commit & Pull Request Guidelines
- Use Conventional Commits: `feat:`, `fix:`, `refactor:`, `docs:`, `test:`, `chore:`.
- Keep PRs focused and small; include a clear summary, linked issues (e.g., `Closes #12`), and local run steps (`cargo run -- <args>`). Add sample output for user-visible changes.
- Before opening a PR, ensure: `cargo fmt --all`, `cargo clippy -- -D warnings`, and `cargo test` all pass.
- Avoid drive‑by refactors/renames. Discuss breaking changes in the PR description.

## Agent‑Specific Instructions
- Scope: This file applies to the entire repository tree.
- Prefer standard library and small, well‑maintained crates; justify new dependencies in the PR description.
- Do not rename files or change the Rust edition without explicit request.
- When adding modules, follow `src/<name>.rs` and update imports (`mod`/`use`) in `lib.rs`/`main.rs` accordingly.

