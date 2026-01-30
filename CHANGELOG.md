# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Initial release of md2docx
- Markdown to DOCX conversion with full formatting support
- Thai/English mixed text with automatic font switching
- Mermaid diagram rendering (pure Rust, no browser required)
- Table of Contents generation with page references
- Custom template support (styles from reference.docx)
- Cover page templates with placeholder substitution
- Headers and footers with chapter names (STYLEREF fields)
- Code blocks with syntax highlighting and line numbers
- Image embedding with captions and cross-references
- Footnotes support
- Cross-references (`{ref:target}` syntax)
- Include directives (`{!include:...}`, `{!code:...}`)
- YAML frontmatter support per chapter
- CLI with `build`, `dump-template`, and `validate-template` commands

### Changed
- N/A

### Fixed
- N/A

## [0.1.0] - 2025-01-30

Initial release.