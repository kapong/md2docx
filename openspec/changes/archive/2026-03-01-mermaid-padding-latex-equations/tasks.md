## 1. Mermaid Diagram Spacing

- [x] 1.1 Add `MermaidSection` struct to `src/config/schema.rs` with `spacing_before` and `spacing_after` fields (default: "120")
- [x] 1.2 Add `mermaid: MermaidSection` field to `ProjectConfig` struct
- [x] 1.3 Pass mermaid spacing config through to `BuildContext` in `src/docx/builder.rs`
- [x] 1.4 Apply `Paragraph::spacing()` to the mermaid image paragraph in `Block::Mermaid` handling in `block_to_elements()`
- [x] 1.5 Add test: mermaid diagram paragraph has default spacing (120/120) when no config
- [x] 1.6 Add test: mermaid diagram paragraph uses custom spacing from config

## 2. Math Configuration

- [x] 2.1 Add `MathSection` struct to `src/config/schema.rs` with `renderer` field (default: "image")
- [x] 2.2 Add `math: MathSection` field to `ProjectConfig` struct
- [x] 2.3 Pass math renderer config through to `BuildContext`

## 3. LaTeX Toolchain Detection

- [x] 3.1 Create `src/docx/math_image.rs` module with toolchain detection (check `latex` and `dvisvgm` in PATH)
- [x] 3.2 Add `mod math_image` to `src/docx/mod.rs`
- [x] 3.3 Implement `is_toolchain_available()` function that caches the result
- [x] 3.4 Add test: detection returns false when executables not found

## 4. LaTeX-to-Image Rendering Pipeline

- [x] 4.1 Implement `render_latex_to_svg(latex: &str) -> Result<Vec<u8>, Error>` in `math_image.rs`
- [x] 4.2 Create minimal LaTeX preamble template for standalone math expressions
- [x] 4.3 Implement temp file management: write `.tex`, run `latex`, run `dvisvgm --exact`
- [x] 4.4 Parse SVG output dimensions for proper DOCX sizing
- [x] 4.5 Implement expression caching (HashMap of LaTeX string → rendered SVG bytes)
- [x] 4.6 Add test: simple expression renders to valid SVG (integration test, requires LaTeX)

## 5. Display Math Image Integration

- [x] 5.1 Update `Block::MathBlock` handling in `block_to_elements()` to check renderer config
- [x] 5.2 When `renderer = "image"`: render LaTeX to SVG, create `DocElement::Image` (centered)
- [x] 5.3 When `renderer = "image"` but toolchain unavailable: fall back to OMML with warning
- [x] 5.4 When `renderer = "omml"`: preserve existing OMML path
- [x] 5.5 Add test: display math produces image element when renderer is "image"
- [x] 5.6 Add test: display math falls back to OMML when toolchain unavailable

## 6. Inline Math Image Integration

- [x] 6.1 Update inline math handling in `inline_to_runs()` to check renderer config
- [x] 6.2 When `renderer = "image"`: render LaTeX to SVG, embed as inline image in run
- [x] 6.3 When `renderer = "image"` but toolchain unavailable: fall back to inline OMML
- [x] 6.4 Add test: inline math produces inline image when renderer is "image"

## 7. Documentation and Cleanup

- [x] 7.1 Document `[mermaid]` config section in `docs/ch05_configuration.md`
- [x] 7.2 Document `[math]` config section in `docs/ch05_configuration.md`
- [x] 7.3 Add LaTeX toolchain installation note to `docs/ch02_installation.md`
- [x] 7.4 Build and verify no compiler warnings
