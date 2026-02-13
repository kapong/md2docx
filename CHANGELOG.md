# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.9] - 2026-02-13

### Fixed

- Code block content no longer modified by image path resolution
- Image width examples now reference existing files
- Table cell text alignment applied correctly via paragraph properties
- Nested blockquotes now have increasing indentation per nesting level
- Soft breaks in blockquotes preserved as line breaks instead of spaces
- Bold and italic formatting now works for Thai (Complex Script) text via w:bCs/w:iCs
- Nested inline formatting (e.g. ***bold+italic***) now renders correctly
- Heading styles now apply bold/italic to Complex Script (Thai) text

## [0.1.8] - 2026-02-12

### Added

- Windows x86_64 binary build support
- PowerShell installer script (install.ps1) for Windows users
- Scoop package manager manifest (bucket/md2docx.json)
- Windows installation documentation in README

### Changed

- Updated install.sh to detect Windows environments and redirect to PowerShell installer
- Release workflow now includes Windows build and Scoop manifest auto-update

## [0.1.7] - 2026-02-11

### Fixed

- Footnote references now display as superscript in body text
- Footnote content at page bottom now shows footnote numbers
- Footnote content uses FootnoteText style without indentation
- No blank lines between footnotes in footer area
- Table header rows now use font size from template instead of default 11pt
- Table layout always uses autofit to contents
- Thai text runs now include `<w:cs/>` and `w:hint="cs"` for proper word wrapping
- Added blank Normal paragraph before section breaks for cleaner chapter endings

### Added

- Thai text samples chapter (ch09) in documentation

## [0.1.0] - 2025-01-30

Initial release.
