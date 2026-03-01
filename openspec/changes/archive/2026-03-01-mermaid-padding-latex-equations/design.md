## Context

The md2docx tool converts Markdown to DOCX. Two rendering areas need improvement:

1. **Mermaid diagrams** are inserted as SVG images via `DocElement::Image` in `builder.rs` (line ~1151). The image paragraph has no explicit spacing — it inherits the default (often 0), making diagrams feel cramped against preceding/following content.

2. **Math equations** use a custom LaTeX-to-OMML converter (`src/docx/math.rs`, ~1100 lines) that translates LaTeX into Word's Office Math Markup Language. While functional, OMML rendering is inconsistent across Word versions and produces poor output for complex expressions (matrices, multi-line, decorations). A LaTeX-to-image pipeline would produce pixel-perfect results.

Current relevant code paths:

- `src/docx/builder.rs`: `Block::Mermaid` → `DocElement::Image`, `Block::MathBlock` → `DocElement::MathBlock(omml)`
- `src/docx/math.rs`: `latex_to_omml_paragraph()`, `latex_to_omml_inline()`
- `src/config/schema.rs`: No mermaid or math configuration sections exist yet

## Goals / Non-Goals

**Goals:**

- Add configurable before/after spacing on mermaid diagram paragraphs
- Implement LaTeX-to-image rendering for display math blocks
- Implement LaTeX-to-image rendering for inline math
- Keep OMML as automatic fallback when image renderer is unavailable
- Make the rendering mode configurable (`image` vs `omml`)

**Non-Goals:**

- Rewriting the OMML converter (it remains as-is for fallback)
- Supporting MathJax/KaTeX server-side rendering (too heavy)
- Editing or improving SVG quality of mermaid output
- Interactive equation editing in Word

## Decisions

### D1: Mermaid spacing — paragraph spacing properties

Add `spacing_before` and `spacing_after` (in twips, like other spacing in the codebase) to the paragraph wrapping mermaid images. Use `w:spacing w:before="..." w:after="..."` on the image paragraph.

**Default**: 120 twips before, 120 twips after (~2mm each — enough breathing room without being excessive).

**Configuration**: Add optional `[mermaid]` section to `md2docx.toml`:

```toml
[mermaid]
spacing_before = "120"
spacing_after = "120"
```

**Rationale**: Paragraph-level spacing is the standard OOXML approach. Using the existing `Paragraph::spacing()` method keeps the change minimal.

**Alternative considered**: Using an empty paragraph before/after — rejected because it creates selectable whitespace and complicates editing.

### D2: LaTeX-to-image — use external TeX engine or embedded tectonic library

Use a LaTeX engine to render LaTeX to SVG images. Supports three engine modes:

1. **Embedded tectonic library** (feature `tectonic-lib`): Uses the tectonic Rust crate compiled into the binary. No separate TeX installation needed (still requires `dvisvgm`).
2. **Tectonic CLI**: Uses externally installed `tectonic` command.
3. **Traditional latex**: Uses `latex` from TeX Live / MacTeX / BasicTeX.

The rendering pipeline:

1. Write the LaTeX expression to a temporary `.tex` file with a minimal preamble (`\documentclass[<font_size>]{article}` with amsmath/amssymb/amsfonts packages)
2. Run TeX engine → XDV/DVI output
3. Run `dvisvgm --exact --no-fonts` → SVG vector image
4. Read the image back and embed it as a `DocElement::Image` (same path as mermaid)

**TeX toolchain PATH discovery (`augmented_path()`)**: Child processes for CLI tools use an augmented `PATH` that appends well-known TeX installation directories (e.g., `/Library/TeX/texbin`, TeX Live year-versioned paths for macOS/Linux). This ensures TeX tools are found even when not in the user's shell PATH.

**LaTeX font size**: Defaults to `10pt` (compact, suitable for document embedding). Configurable via `[math] font_size` in md2docx.toml. Valid values: `"8pt"`, `"9pt"`, `"10pt"`, `"11pt"`, `"12pt"`. Invalid values fall back to `10pt`.

**Rendering format**: SVG preferred (vector, crisp at any zoom). PNG as fallback if SVG pipeline unavailable.

**Rationale**: A real LaTeX engine produces the best output. The mermaid rendering already proves the SVG-in-DOCX path works well.

**Alternative considered**:

- Pure-Rust math renderer (e.g., `rex`) — too limited, poor coverage of LaTeX commands
- KaTeX CLI (Node.js) — extra runtime dependency, but could be a future option
- Keep OMML only — rejected because quality is the core problem

### D3: Math rendering mode configuration

Add `[math]` section to `md2docx.toml`:

```toml
[math]
renderer = "image"    # "image" | "auto" | "omml" (default: "image")
font_size = "10pt"    # "8pt" | "9pt" | "10pt" | "11pt" | "12pt" (default: "10pt")
```

When `renderer = "image"`:

- Display math (`$$...$$`) → render to SVG, embed as centered image
- Inline math (`$...$`) → render to SVG, embed as inline image (anchored in text run)
- If LaTeX toolchain unavailable → fall back to OMML with a warning

When `renderer = "auto"`:

- Detect toolchain at startup; use "image" if LaTeX+dvisvgm are available, "omml" otherwise
- Provides zero-config experience: best quality when tools are present, graceful fallback when not

When `renderer = "omml"`:

- Use existing OMML path (current behavior)

**Rationale**: Allows users without LaTeX installed to still get math output. Makes the feature opt-out rather than breaking existing setups.

### D4: Image sizing for equations

- Display math: auto-sized from SVG viewBox, constrained to page width
- Inline math: auto-sized from SVG, vertically aligned with text baseline using `w:drawing` inline properties
- Use tight cropping in the LaTeX→SVG pipeline (dvisvgm `--exact` flag)

### D5: Module structure

- New module: `src/docx/math_image.rs` — handles LaTeX-to-image rendering
- `src/docx/math.rs` — unchanged (OMML converter, used as fallback)
- `src/docx/builder.rs` — updated to choose rendering path based on config
- `src/config/schema.rs` — new `MermaidSection` and `MathSection` structs

## Risks / Trade-offs

- **[External dependency on LaTeX toolchain]** → Mitigation: OMML fallback when `latex`/`dvisvgm` not found. Clear error message guiding installation. Document in README.
- **[Performance: spawning external processes per equation]** → Mitigation: batch equations, cache rendered images for identical expressions within a document.
- **[Inline math vertical alignment]** → Mitigation: use SVG baseline metadata from dvisvgm; may need manual offset tuning. Acceptable for v1 — can refine later.
- **[Cross-platform LaTeX availability]** → Mitigation: TeX Live available on all platforms. Docker image already has it. Document minimal install (e.g., `basictex` on macOS).
