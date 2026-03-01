## Context

The md2docx tool renders Mermaid diagrams using `mermaid-rs-renderer`, a pure-Rust library that only supports graph/flowchart layouts. The SVG output is embedded directly in the DOCX via `DocElement::Image`. The `mermaid-png` feature (enabled by default) already provides `render_to_png()` and `svg_to_png()` functions using `resvg`/`tiny-skia`, but the builder currently only calls `render_to_svg()`.

Current code path in `src/docx/builder.rs` (Block::Mermaid handler):
1. `crate::mermaid::render_to_svg(content)` → SVG string
2. `ctx.image_ctx.add_image_data(filename.svg, svg_bytes, ...)` → relationship ID
3. Dimension extraction from image context
4. `ImageElement::new(...)` with mermaid spacing applied

The `render_to_png(content, scale)` function already exists and works: it calls `render_to_svg()` internally, then converts via `svg_to_png()` using resvg. The scale factor controls resolution (e.g., scale=2.0 at base ~75 DPI ≈ 150 DPI).

`MermaidSection` in `src/config/schema.rs` currently has only `spacing_before` and `spacing_after` fields.

## Goals / Non-Goals

**Goals:**

- Switch default mermaid output from SVG to PNG at 150 DPI
- Add `output_format` config field to `[mermaid]` section (`"png"` default, `"svg"` opt-in)
- Add `dpi` config field to `[mermaid]` section (default: `150`)
- Maintain identical visual output quality
- Keep SVG as an opt-in alternative

**Non-Goals:**

- Adding external `mmdc` CLI support (future work)
- Supporting diagram types beyond graph/flowchart (renderer limitation, not output format)
- Changing math equation rendering (separate concern)

## Decisions

### D1: PNG as default output format

Switch the `Block::Mermaid` handler from `render_to_svg()` to `render_to_png()`. The PNG path already exists and is feature-gated behind `mermaid-png` (enabled by default).

**When `mermaid-png` feature is not enabled**: Fall back to SVG (current behavior). Emit no warning since SVG still works.

**Rationale**: PNG is universally supported in all Word versions. SVG support varies and can produce rendering artifacts.

### D2: DPI calculation via scale factor

The existing `render_to_png(content, scale)` takes a scale factor, not DPI directly. The base SVG rendering produces output at roughly 75 DPI (standard screen resolution for SVG viewBox units). To achieve the user-configured DPI:

```
scale = dpi / 75.0
```

For 150 DPI: `scale = 150.0 / 75.0 = 2.0`
For 300 DPI: `scale = 300.0 / 75.0 = 4.0`

**EMU dimension calculation**: After rendering to PNG, the DOCX image dimensions must be calculated from the PNG pixel dimensions and the target DPI:

```
width_emu = (png_width_px / dpi) * 914400
height_emu = (png_height_px / dpi) * 914400
```

This ensures the image appears at the correct physical size in Word regardless of DPI.

**Rationale**: Higher DPI renders cleaner text/lines but increases file size. 150 DPI is a reasonable balance (4x pixels vs 75 DPI, ~4x file size increase per image). Users who need print quality can set 300 DPI.

### D3: Configuration structure

Add two new fields to `MermaidSection`:

```toml
[mermaid]
output_format = "png"     # "png" (default) or "svg"
dpi = 150                 # PNG resolution (default: 150, ignored for SVG)
spacing_before = "120"
spacing_after = "120"
```

`output_format` is a string field. `dpi` is a `u32` field.

**Rationale**: Simple flat config. DPI as integer avoids floating-point confusion. The DPI field is ignored when `output_format = "svg"`.

### D4: Filename extension

When outputting PNG, use `.png` extension in the virtual filename (`mermaid{id}.png` instead of `mermaid{id}.svg`). The `add_image_data()` function uses the filename extension to determine the OOXML content type for the relationship.

### D5: Fallback when feature not available

When `output_format = "png"` but the `mermaid-png` feature is not compiled in, fall back to SVG silently. The `render_to_png()` stub already returns an error in this case — catch it and retry with `render_to_svg()`.

## Risks / Trade-offs

- **[File size increase]**: PNG at 150 DPI produces larger images than SVG (~4-10x per diagram). Acceptable for document use. Users can reduce with lower DPI or opt into SVG.
- **[Quality at low DPI]**: Below 150 DPI, text in diagrams may look blurry. Mitigated by making DPI configurable with a sensible default.
- **[Breaking change for existing users]**: Output format changes from SVG to PNG. Functionally equivalent but file internals differ. Not a user-visible breaking change since Word renders both formats. Users who specifically depend on SVG can opt in via config.
