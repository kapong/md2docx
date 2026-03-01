## Why

Mermaid diagrams are currently rendered as SVG and embedded directly in DOCX. While SVG works for simple flowcharts, many Word versions (especially older ones and Word Online) have limited or broken SVG rendering. Additionally, the pure-Rust mermaid renderer (`mermaid-rs-renderer`) only supports graph/flowchart diagrams — unsupported types (sequence, ER, class, state, pie, etc.) are skipped entirely. Switching to PNG output at 150 DPI as the default format ensures consistent rendering across all Word versions and opens the door to supporting all diagram types via the external `mmdc` CLI fallback.

## What Changes

- **Change default mermaid output from SVG to PNG (150 DPI)**: The `Block::Mermaid` handler in the builder will call `render_to_png()` instead of `render_to_svg()`, producing raster images at 150 DPI for reliable Word compatibility
- **Add `output_format` configuration**: New `output_format` field in `[mermaid]` config (`"png"` default, `"svg"` available) so users can opt back into SVG if desired
- **Add `dpi` configuration**: New `dpi` field in `[mermaid]` config (default: `150`) to control PNG resolution
- **Enable `mermaid-png` feature by default**: The feature is already in the default feature set; ensure the builder always uses the PNG path when available

## Capabilities

### New Capabilities

- `mermaid-png-output`: Mermaid diagrams rendered as PNG images at configurable DPI instead of SVG

### Modified Capabilities

- `mermaid-spacing`: The spacing behavior remains identical but applies to PNG image paragraphs instead of SVG

## Impact

- **Code**: `src/docx/builder.rs` (Block::Mermaid handler switches from `render_to_svg` to `render_to_png`), `src/config/schema.rs` (new fields on `MermaidSection`), `src/mermaid/mod.rs` (DPI-aware PNG rendering)
- **Dependencies**: `resvg`, `usvg`, `tiny-skia` already present via `mermaid-png` feature — no new dependencies
- **Configuration**: `[mermaid]` section gains `output_format` and `dpi` fields
- **Output**: DOCX files will contain PNG images instead of SVG for mermaid diagrams (visual output should be identical or better)
- **Existing behavior**: Users can set `output_format = "svg"` to preserve current SVG behavior
