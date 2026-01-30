# Contributing to md2docx

Thank you for your interest in contributing to md2docx!

## Getting Started

1. Fork the repository
2. Clone your fork: `git clone https://github.com/YOUR_USERNAME/md2docx.git`
3. Create a branch: `git checkout -b feature/your-feature-name`

## Development Setup

### Prerequisites
- Rust 1.75+ (install via [rustup](https://rustup.rs/))
- Git

### Build
```bash
cargo build
```

### Test
```bash
cargo test
```

### Lint
```bash
cargo clippy
cargo fmt --check
```

## Pull Request Guidelines

1. **Create an issue first** - Discuss major changes before implementing
2. **Keep PRs focused** - One feature or fix per PR
3. **Write tests** - Add tests for new functionality
4. **Update documentation** - Keep AGENTS.md and README updated
5. **Follow conventional commits**:
   - `feat:` - New features
   - `fix:` - Bug fixes
   - `refactor:` - Code refactoring
   - `docs:` - Documentation changes
   - `test:` - Test additions/changes
   - `chore:` - Maintenance tasks

## Code Style

- Run `cargo fmt` before committing
- Run `cargo clippy` and fix all warnings
- No `unwrap()` in library code (use `?` or `expect()` with message)
- Add doc comments to public APIs

## Testing

- Unit tests go in the same file as the code
- Integration tests go in `tests/` directory
- Run all tests: `cargo test`

## License

By contributing, you agree that your contributions will be licensed under the MIT License.