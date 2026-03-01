## Why

Mermaid diagrams render without vertical spacing, causing them to sit flush against surrounding text — making the document feel cramped. Additionally, the current Word-native OMML equation rendering produces inconsistent and poor-quality output for complex LaTeX expressions. Switching to a LaTeX-to-image pipeline would produce high-fidelity equation images that render identically across all Word versions.

## What Changes

- Add configurable vertical padding (spacing before/after) around mermaid diagram paragraphs so diagrams have breathing room in the final DOCX
- Replace the OMML-based math block rendering path with a LaTeX-to-image approach: render LaTeX expressions to images (PNG/SVG) and embed them as inline or block images in the DOCX
- Keep OMML as a fallback when the image renderer is unavailable
- Inline math (`$...$`) will also use the image path when the renderer is available, falling back to OMML

## Capabilities

### New Capabilities

- `mermaid-spacing`: Configurable spacing (before/after) for mermaid diagram paragraphs in DOCX output
- `latex-equation-image`: Render LaTeX math expressions to images and embed them in DOCX instead of using Word-native OMML

### Modified Capabilities
<!-- No existing spec-level requirement changes -->

## Impact

- **Code**: `src/docx/builder.rs` (mermaid block rendering, math block rendering), `src/docx/math.rs` (may become secondary/fallback), new module for LaTeX-to-image rendering
- **Dependencies**: Will need a LaTeX-to-image rendering library or external tool (e.g., `mathjax`, `KaTeX` via CLI, or a Rust crate)
- **Configuration**: `src/config/schema.rs` — new config fields for mermaid spacing values and equation rendering mode
- **Templates**: Image template settings may apply to equation images (sizing, alignment)
- **Existing behavior**: OMML rendering remains as fallback; no breaking changes
