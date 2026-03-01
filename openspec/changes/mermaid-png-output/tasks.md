## 1. Configuration

- [ ] 1.1 Add `output_format: String` field to `MermaidSection` in `src/config/schema.rs` (default: `"png"`)
- [ ] 1.2 Add `dpi: u32` field to `MermaidSection` in `src/config/schema.rs` (default: `150`)
- [ ] 1.3 Add `mermaid_output_format: String` and `mermaid_dpi: u32` fields to `DocumentConfig` in `src/docx/builder.rs`
- [ ] 1.4 Pass mermaid output format and DPI through from `ProjectConfig` to `DocumentConfig` in `src/project/mod.rs`
- [ ] 1.5 Add `mermaid_output_format` and `mermaid_dpi` to `BuildContextParams` and `BuildContext`

## 2. Builder PNG Integration

- [ ] 2.1 Update `Block::Mermaid` handler in `block_to_elements()` to check `ctx.mermaid_output_format`
- [ ] 2.2 When `output_format = "png"`: call `render_to_png(content, scale)` with `scale = dpi / 75.0`
- [ ] 2.3 When `output_format = "png"`: use `.png` extension in virtual filename
- [ ] 2.4 When `output_format = "png"`: calculate EMU dimensions from PNG pixel size and DPI
- [ ] 2.5 When `output_format = "svg"`: preserve existing SVG rendering path
- [ ] 2.6 When `output_format = "png"` but `render_to_png` fails (feature not compiled): fall back to SVG silently

## 3. Nested Context (BlockQuote Mermaid)

- [ ] 3.1 Update mermaid handling in `blockquote_to_elements()` to use same PNG/SVG logic

## 4. Tests

- [ ] 4.1 Add test: default `MermaidSection` has `output_format = "png"` and `dpi = 150`
- [ ] 4.2 Add test: `DocumentConfig` default has `mermaid_output_format = "png"` and `mermaid_dpi = 150`
- [ ] 4.3 Add test: mermaid PNG rendering produces valid PNG bytes (integration, requires `mermaid-png` feature)
- [ ] 4.4 Add test: EMU dimensions calculated correctly from PNG pixels and DPI

## 5. Documentation

- [ ] 5.1 Document `output_format` and `dpi` fields in `[mermaid]` section of `docs/ch05_configuration.md`
- [ ] 5.2 Build and verify no compiler warnings
