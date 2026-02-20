//! Pure-Rust LaTeX math renderer using ReX (no external tools needed)
//!
//! Uses the ReX mathematical typesetting engine with a custom SVG backend
//! to render LaTeX math expressions to SVG without any external dependencies.

use rex::font::backend::ttf_parser::TtfMathFont;
use rex::font::common::GlyphId;
use rex::layout::engine::LayoutBuilder;
use rex::render::{Backend, Cursor, FontBackend, GraphicsBackend, RGBA, Renderer, Role};

use std::collections::HashMap;
use std::sync::Mutex;
use once_cell::sync::Lazy;

use crate::error::Error;

/// Embedded XITS Math font (OpenType math font)
static MATH_FONT_DATA: &[u8] = include_bytes!("XITS_Math.otf");

// ── SVG Backend ────────────────────────────────────────────────────────

/// A pure SVG backend for ReX that generates SVG path data directly
/// from font glyph outlines, without requiring Cairo or any C libraries.
struct SvgBackend<'a> {
    /// Accumulated SVG elements
    elements: Vec<String>,
    /// Current fill color
    current_color: (u8, u8, u8, u8),
    /// Color stack for begin_color/end_color
    color_stack: Vec<(u8, u8, u8, u8)>,
    /// Reference to the font for glyph outline extraction
    #[allow(dead_code)]
    font: &'a TtfMathFont<'a>,
}

impl<'a> SvgBackend<'a> {
    fn new(font: &'a TtfMathFont<'a>) -> Self {
        Self {
            elements: Vec::new(),
            current_color: (0, 0, 0, 255),
            color_stack: Vec::new(),
            font,
        }
    }

    fn color_str(&self) -> String {
        let (r, g, b, _a) = self.current_color;
        if r == 0 && g == 0 && b == 0 {
            "black".to_string()
        } else {
            format!("rgb({},{},{})", r, g, b)
        }
    }

    fn into_svg(self, width: f64, height: f64, x_min: f64, y_min: f64) -> String {
        let mut svg = String::with_capacity(4096);
        // Use width/height for the viewBox but add small padding
        let pad = 1.0;
        svg.push_str(&format!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" width="{:.1}" height="{:.1}" viewBox="{:.2} {:.2} {:.2} {:.2}">"#,
            width + pad * 2.0,
            height + pad * 2.0,
            x_min - pad,
            y_min - pad,
            width + pad * 2.0,
            height + pad * 2.0,
        ));
        svg.push('\n');
        for elem in &self.elements {
            svg.push_str(elem);
            svg.push('\n');
        }
        svg.push_str("</svg>");
        svg
    }
}

/// SVG path builder that implements ttf_parser::OutlineBuilder
struct SvgPathBuilder {
    path: String,
}

impl SvgPathBuilder {
    fn new() -> Self {
        Self {
            path: String::with_capacity(256),
        }
    }
}

impl ttf_parser::OutlineBuilder for SvgPathBuilder {
    fn move_to(&mut self, x: f32, y: f32) {
        self.path.push_str(&format!("M{:.2} {:.2}", x, y));
    }

    fn line_to(&mut self, x: f32, y: f32) {
        self.path.push_str(&format!("L{:.2} {:.2}", x, y));
    }

    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.path.push_str(&format!("Q{:.2} {:.2} {:.2} {:.2}", x1, y1, x, y));
    }

    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.path.push_str(&format!(
            "C{:.2} {:.2} {:.2} {:.2} {:.2} {:.2}",
            x1, y1, x2, y2, x, y
        ));
    }

    fn close(&mut self) {
        self.path.push('Z');
    }
}

impl<'a> GraphicsBackend for SvgBackend<'a> {
    fn bbox(&mut self, _pos: Cursor, _width: f64, _height: f64, _role: Role) {
        // Debug bounding boxes - skip in production
    }

    fn rule(&mut self, pos: Cursor, width: f64, height: f64) {
        let color = self.color_str();
        self.elements.push(format!(
            r#"<rect x="{:.2}" y="{:.2}" width="{:.2}" height="{:.2}" fill="{}"/>"#,
            pos.x, pos.y, width, height, color
        ));
    }

    fn begin_color(&mut self, color: RGBA) {
        let old = std::mem::replace(
            &mut self.current_color,
            (color.0, color.1, color.2, color.3),
        );
        self.color_stack.push(old);
    }

    fn end_color(&mut self) {
        if let Some(color) = self.color_stack.pop() {
            self.current_color = color;
        }
    }
}

impl<'a> FontBackend<TtfMathFont<'a>> for SvgBackend<'a> {
    fn symbol(&mut self, pos: Cursor, gid: GlyphId, scale: f64, ctx: &TtfMathFont<'a>) {
        let mut builder = SvgPathBuilder::new();

        // Extract glyph outline
        ctx.font().outline_glyph(gid.into(), &mut builder);

        if builder.path.is_empty() {
            return;
        }

        let color = self.color_str();
        let sx = scale * f64::from(ctx.font_matrix().sx);
        let sy = -scale * f64::from(ctx.font_matrix().sy); // flip Y axis

        self.elements.push(format!(
            r#"<path d="{}" fill="{}" fill-rule="evenodd" transform="translate({:.2},{:.2}) scale({:.6},{:.6})"/>"#,
            builder.path, color, pos.x, pos.y, sx, sy
        ));
    }
}

impl<'a> Backend<TtfMathFont<'a>> for SvgBackend<'a> {}

// ── LaTeX Preprocessing ────────────────────────────────────────────────

/// Preprocess LaTeX to handle constructs that need workarounds.
///
/// `\sqrt{...}` and `\sqrt[2]{...}` are passed through as-is (ReX handles them natively).
/// For any other `\sqrt[n]{...}` (n ≠ 2), rewrites to `{}^{n}\!\sqrt{...}` to render
/// the degree as a superscript prefix.
fn preprocess_latex(latex: &str) -> String {
    let mut result = String::with_capacity(latex.len());
    let chars: Vec<char> = latex.chars().collect();
    let len = chars.len();
    let mut i = 0;

    while i < len {
        // Look for \sqrt[
        if i + 5 < len && &latex[char_offset(&chars, i)..char_offset(&chars, i + 5)] == r"\sqrt" {
            if i + 5 < len && chars[i + 5] == '[' {
                // Found \sqrt[ — extract the index
                if let Some(bracket_end) = find_closing_bracket(&chars, i + 5) {
                    let index: String = chars[i + 6..bracket_end].iter().collect();
                    let index = index.trim();
                    if index == "2" {
                        // \sqrt[2]{...} → \sqrt{...} (square root, skip redundant index)
                        result.push_str(r"\sqrt");
                    } else {
                        // nth root (n ≠ 2): {}^{n}\!\sqrt
                        result.push_str(&format!("{{}}^{{{}}}\\!\\sqrt", index));
                    }
                    i = bracket_end + 1; // skip past ']'
                    continue;
                }
            }
            // Plain \sqrt (no optional arg) — copy as-is
            result.push_str(r"\sqrt");
            i += 5;
            continue;
        }
        result.push(chars[i]);
        i += 1;
    }
    result
}

/// Get byte offset for a char index
fn char_offset(chars: &[char], idx: usize) -> usize {
    chars[..idx].iter().map(|c| c.len_utf8()).sum()
}

/// Find closing ']' for an opening '[' at position `start`
fn find_closing_bracket(chars: &[char], start: usize) -> Option<usize> {
    debug_assert_eq!(chars[start], '[');
    let mut depth = 1;
    let mut i = start + 1;
    while i < chars.len() {
        match chars[i] {
            '[' => depth += 1,
            ']' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
            }
            _ => {}
        }
        i += 1;
    }
    None
}

/// Check if a LaTeX string contains non-Latin/non-math characters inside `\text{...}`
/// blocks that cannot be rendered by the XITS Math font.
///
/// Returns `Some(char)` with the first unsupported character found, or `None`.
fn find_unsupported_text_char(latex: &str) -> Option<char> {
    // Find all \text{...} blocks and check their content
    let bytes = latex.as_bytes();
    let len = bytes.len();
    let mut i = 0;

    while i < len {
        // Look for \text{ pattern
        if i + 6 < len && &latex[i..i + 5] == r"\text" {
            // Skip to the opening brace
            let mut j = i + 5;
            while j < len && bytes[j] == b' ' {
                j += 1;
            }
            if j < len && bytes[j] == b'{' {
                // Find matching close brace (handling nesting)
                let mut depth = 1;
                let start = j + 1;
                j += 1;
                while j < len && depth > 0 {
                    if bytes[j] == b'{' {
                        depth += 1;
                    } else if bytes[j] == b'}' {
                        depth -= 1;
                    }
                    if depth > 0 {
                        j += 1;
                    }
                }
                // Check chars in \text{...} content
                let text_content = &latex[start..j];
                for ch in text_content.chars() {
                    // Allow ASCII, common punctuation, and basic Latin extensions
                    if !ch.is_ascii() && !is_math_font_supported(ch) {
                        return Some(ch);
                    }
                }
                i = j + 1;
                continue;
            }
        }
        i += 1;
    }
    None
}

/// Check if a non-ASCII character is likely supported by the XITS Math font.
/// This covers Latin Extended, Greek, Cyrillic, and mathematical symbol ranges.
fn is_math_font_supported(ch: char) -> bool {
    let cp = ch as u32;
    matches!(cp,
        0x00A0..=0x024F  // Latin Extended-A/B
        | 0x0370..=0x03FF // Greek and Coptic
        | 0x0400..=0x04FF // Cyrillic
        | 0x1D00..=0x1DBF // Phonetic Extensions
        | 0x2000..=0x206F // General Punctuation
        | 0x2070..=0x209F // Superscripts and Subscripts
        | 0x20A0..=0x20CF // Currency Symbols 
        | 0x20D0..=0x20FF // Combining Diacritical Marks for Symbols
        | 0x2100..=0x214F // Letterlike Symbols
        | 0x2150..=0x218F // Number Forms
        | 0x2190..=0x21FF // Arrows
        | 0x2200..=0x22FF // Mathematical Operators
        | 0x2300..=0x23FF // Miscellaneous Technical
        | 0x2460..=0x24FF // Enclosed Alphanumerics
        | 0x2500..=0x257F // Box Drawing
        | 0x25A0..=0x25FF // Geometric Shapes
        | 0x2600..=0x26FF // Miscellaneous Symbols
        | 0x2700..=0x27BF // Dingbats
        | 0x27C0..=0x27EF // Miscellaneous Mathematical Symbols-A
        | 0x27F0..=0x27FF // Supplemental Arrows-A
        | 0x2900..=0x297F // Supplemental Arrows-B
        | 0x2980..=0x29FF // Miscellaneous Mathematical Symbols-B
        | 0x2A00..=0x2AFF // Supplemental Mathematical Operators
        | 0x2B00..=0x2BFF // Miscellaneous Symbols and Arrows
        | 0x1D400..=0x1D7FF // Mathematical Alphanumeric Symbols
        | 0xFB00..=0xFB06 // Alphabetic Presentation Forms (ligatures)
    )
}

// ── Public API ─────────────────────────────────────────────────────────

/// Cache for rendered math expressions
static RENDER_CACHE: Lazy<Mutex<HashMap<String, MathSvgResult>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Rich metadata from a rendered math SVG, providing everything needed
/// for correct sizing and vertical alignment without ad-hoc heuristics.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct MathSvgResult {
    /// The SVG image bytes
    pub svg_bytes: Vec<u8>,
    /// Total width in EMU
    pub width_emu: i64,
    /// Total height in EMU (ascent + descent)
    pub height_emu: i64,
    /// Distance from baseline to top of the formula, in EMU (positive)
    pub ascent_emu: i64,
    /// Distance from baseline to bottom of the formula, in EMU (positive, measures downward)
    pub descent_emu: i64,
    /// The font size used for rendering, in EMU
    pub font_size_emu: i64,
    /// Pixels-per-inch used during rendering (96 for SVG)
    pub ppi: f64,
    /// Suggested `w:position` value in half-points for OOXML vertical alignment.
    /// Negative = lower. For inline math this aligns the formula baseline with
    /// the surrounding text baseline. For display math, centres tall formulas
    /// on the text line. `None` when no adjustment is needed.
    pub position: Option<i32>,
}

/// Render a LaTeX math expression to SVG bytes using ReX.
///
/// Returns a [`MathSvgResult`] with SVG data, dimensions in EMU, baseline
/// metrics (ascent / descent), and a pre-computed `w:position` value that
/// consuming code can apply directly to `ImageElement.position`.
///
/// This is a pure-Rust implementation that requires no external tools.
pub fn render_latex_to_svg(
    latex: &str,
    display: bool,
    font_size_str: &str,
) -> Result<MathSvgResult, Error> {
    // Check cache
    let cache_key = format!("rex:{}:{}:{}", latex, display, font_size_str);
    if let Ok(cache) = RENDER_CACHE.lock() {
        if let Some(cached) = cache.get(&cache_key) {
            return Ok(cached.clone());
        }
    }

    // Parse font size (e.g. "10pt" -> 10)
    let font_size_pt: f64 = font_size_str
        .trim_end_matches("pt")
        .parse::<f64>()
        .unwrap_or(10.0)
        .clamp(8.0, 24.0);

    // XITS Math font renders ~20% larger than Computer Modern (LaTeX default).
    // Scale down to match tectonic/dvisvgm output visually.
    const XITS_SCALE: f64 = 0.80;
    let font_size_px = font_size_pt * (96.0 / 72.0) * XITS_SCALE;

    // Load font
    let face = ttf_parser::Face::parse(MATH_FONT_DATA, 0)
        .map_err(|e| Error::Math(format!("Failed to parse math font: {}", e)))?;
    let math_font = TtfMathFont::new(face)
        .map_err(|e| Error::Math(format!("Font lacks MATH table: {:?}", e)))?;

    // Preprocess LaTeX for ReX compatibility (e.g. \sqrt[n]{...})
    let latex = preprocess_latex(latex);
    let latex = latex.as_str();

    // Check for non-Latin characters in \text{} blocks that the math font cannot render
    if let Some(ch) = find_unsupported_text_char(latex) {
        return Err(Error::Math(format!(
            "\\text{{}} contains non-Latin character '{}' (U+{:04X}) unsupported by math font",
            ch, ch as u32
        )));
    }

    // Parse and layout
    let layout_engine = LayoutBuilder::new(&math_font)
        .font_size(font_size_px)
        .build();

    let parse_nodes = rex::parser::parse(latex)
        .map_err(|e| Error::Math(format!("LaTeX parse error: {:?}", e)))?;

    let layout = layout_engine
        .layout(&parse_nodes)
        .map_err(|e| Error::Math(format!("Layout error: {:?}", e)))?;

    // Get dimensions from the full bounding box (for SVG rendering)
    let mut bbox = layout.full_bounding_box();
    let width = bbox.width();
    let mut height = bbox.height();

    if width <= 0.0 || height <= 0.0 {
        return Err(Error::Math("Layout produced zero-size output".to_string()));
    }

    // Ensure the SVG spans at least the font em-square height for display math.
    // Without this, short glyphs get a tiny bounding box. We pad the viewBox
    // (adding whitespace below baseline) instead of scaling the glyph.
    // For inline math, skip this — the natural bounding box is correct and
    // Word handles baseline alignment via w:position.
    if display && height < font_size_px {
        let deficit = font_size_px - height;
        bbox.y_min -= deficit; // extend below baseline
        height = font_size_px;
    }

    // Get typographic metrics from layout: baseline-relative ascent & descent
    let dims = layout.size();
    // dims.height = distance baseline → top  (positive)
    // dims.depth  = distance baseline → bottom (negative, so negate for positive descent)
    let ascent_px = dims.height.abs();
    let descent_px = dims.depth.abs();

    // Render to SVG
    let mut backend = SvgBackend::new(&math_font);
    let renderer = Renderer::new();
    renderer.render(&layout, &mut backend);

    let svg_string = backend.into_svg(width, height, bbox.x_min, bbox.y_min);
    let svg_bytes = svg_string.into_bytes();

    // Convert dimensions to EMU (English Metric Units)
    // 1 inch = 914400 EMU, 1 px at 96dpi = 914400/96 = 9525 EMU
    // into_svg adds 1px padding on each side, so EMU must match the SVG's
    // intrinsic size (content + 2 * pad) — otherwise Word squeezes the image.
    const PPI: f64 = 96.0;
    const SVG_PAD: f64 = 1.0;
    let emu_per_px: f64 = 914400.0 / PPI;
    let width_emu = ((width + SVG_PAD * 2.0) * emu_per_px) as i64;
    let height_emu = ((height + SVG_PAD * 2.0) * emu_per_px) as i64;
    let ascent_emu = (ascent_px * emu_per_px) as i64;
    let descent_emu = (descent_px * emu_per_px) as i64;
    // font_size_emu reflects the actual rendered size (including XITS_SCALE)
    let font_size_emu = (font_size_px * emu_per_px) as i64;

    // Compute w:position (half-points, negative = lower) for OOXML vertical alignment.
    // 1 half-point = 6350 EMU.
    //
    // Word places inline images with their bottom edge on the text baseline.
    // The math formula's baseline sits at (descent + SVG_PAD) pixels above
    // the bottom of the SVG image. We shift the image down by that amount
    // so the formula baseline aligns with the text baseline.
    //
    // For display math, centre tall equations on the text line instead.
    let pad_emu = (SVG_PAD * emu_per_px) as i64;
    let position = if !display {
        // Inline math: lower by (descent + pad) to align baselines
        let offset_emu = (descent_emu + pad_emu).max(1);
        Some(-((offset_emu as f64 / 6350.0).round() as i32).max(1))
    } else {
        // Display math: centre tall equations
        if height_emu > font_size_emu {
            let offset_emu = (height_emu - font_size_emu) / 2;
            Some(-((offset_emu as f64 / 6350.0).round() as i32).max(1))
        } else {
            None
        }
    };

    let result = MathSvgResult {
        svg_bytes,
        width_emu,
        height_emu,
        ascent_emu,
        descent_emu,
        font_size_emu,
        ppi: PPI,
        position,
    };

    // Cache the result
    if let Ok(mut cache) = RENDER_CACHE.lock() {
        cache.insert(cache_key, result.clone());
    }

    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_expression() {
        let result = render_latex_to_svg("x + y", false, "10pt");
        assert!(result.is_ok(), "Failed: {:?}", result.err());
        let math = result.unwrap();
        assert!(math.width_emu > 0);
        assert!(math.height_emu > 0);
        assert!(math.ascent_emu > 0);
        assert!(math.ppi == 96.0);
        let svg_str = String::from_utf8(math.svg_bytes).unwrap();
        assert!(svg_str.starts_with("<svg"));
        assert!(svg_str.contains("<path"));
    }

    #[test]
    fn test_fraction() {
        let result = render_latex_to_svg(r"\frac{a}{b}", true, "10pt");
        assert!(result.is_ok(), "Failed: {:?}", result.err());
        let math = result.unwrap();
        assert!(math.width_emu > 0);
        assert!(math.height_emu > 0);
        // Fractions should have both ascent and descent
        assert!(math.ascent_emu > 0);
        assert!(math.descent_emu > 0);
        // Display math position should be set for tall equations
        assert!(math.position.is_some());
    }

    #[test]
    fn test_complex_equation() {
        let result = render_latex_to_svg(
            r"E = mc^2",
            true,
            "10pt",
        );
        assert!(result.is_ok(), "Failed: {:?}", result.err());
    }

    #[test]
    fn test_integral() {
        let result = render_latex_to_svg(
            r"\int_0^1 x^2 \, dx",
            true,
            "10pt",
        );
        assert!(result.is_ok(), "Failed: {:?}", result.err());
    }

    #[test]
    fn test_matrix() {
        let result = render_latex_to_svg(
            r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}",
            true,
            "10pt",
        );
        assert!(result.is_ok(), "Failed: {:?}", result.err());
    }

    #[test]
    fn test_greek_letters() {
        let result = render_latex_to_svg(
            r"\alpha + \beta + \gamma = \pi",
            true,
            "10pt",
        );
        assert!(result.is_ok(), "Failed: {:?}", result.err());
    }

    #[test]
    fn test_font_size_12pt() {
        let result_10 = render_latex_to_svg("x", false, "10pt").unwrap();
        let result_12 = render_latex_to_svg("x", false, "12pt").unwrap();
        // 12pt should produce larger output than 10pt
        assert!(result_12.width_emu > result_10.width_emu, "12pt should be wider than 10pt");
    }

    #[test]
    fn test_cache() {
        let r1 = render_latex_to_svg("a+b", false, "10pt").unwrap();
        let r2 = render_latex_to_svg("a+b", false, "10pt").unwrap();
        assert_eq!(r1.width_emu, r2.width_emu);
        assert_eq!(r1.height_emu, r2.height_emu);
    }

    #[test]
    fn test_cauchy_schwarz() {
        let result = render_latex_to_svg(
            r"\left\vert \sum_k a_kb_k \right\vert \leq \left(\sum_k a_k^2\right)^{\frac12}\left(\sum_k b_k^2\right)^{\frac12}",
            true,
            "10pt",
        );
        assert!(result.is_ok(), "Failed: {:?}", result.err());
    }

    #[test]
    fn test_normal_distribution() {
        let result = render_latex_to_svg(
            r"f(x) = \frac{1}{\sigma\sqrt{2\pi}} e^{-\frac{1}{2}\left(\frac{x-\mu}{\sigma}\right)^2}",
            true,
            "10pt",
        );
        assert!(result.is_ok(), "Failed: {:?}", result.err());
    }

    // ── Preprocessing tests ────────────────────────────────────────────

    #[test]
    fn test_preprocess_sqrt_cuberoot() {
        // n=3 now uses the general workaround
        assert_eq!(preprocess_latex(r"\sqrt[3]{27}"), r"{}^{3}\!\sqrt{27}");
    }

    #[test]
    fn test_preprocess_sqrt_fourthroot() {
        // n=4 now uses the general workaround
        assert_eq!(preprocess_latex(r"\sqrt[4]{x}"), r"{}^{4}\!\sqrt{x}");
    }

    #[test]
    fn test_preprocess_sqrt_nth_root() {
        assert_eq!(
            preprocess_latex(r"\sqrt[5]{32}"),
            r"{}^{5}\!\sqrt{32}"
        );
    }

    #[test]
    fn test_preprocess_plain_sqrt_unchanged() {
        assert_eq!(preprocess_latex(r"\sqrt{x}"), r"\sqrt{x}");
    }

    #[test]
    fn test_preprocess_sqrt2_becomes_plain() {
        // \sqrt[2]{x} is redundant → becomes \sqrt{x}
        assert_eq!(preprocess_latex(r"\sqrt[2]{x}"), r"\sqrt{x}");
    }

    #[test]
    fn test_preprocess_mixed() {
        assert_eq!(
            preprocess_latex(r"x = \sqrt[3]{27} + \sqrt{4}"),
            r"x = {}^{3}\!\sqrt{27} + \sqrt{4}"
        );
    }

    #[test]
    fn test_cuberoot_renders() {
        let result = render_latex_to_svg(r"\sqrt[3]{27} = 3", true, "10pt");
        assert!(result.is_ok(), "cube root failed: {:?}", result.err());
        let math = result.unwrap();
        assert!(math.width_emu > 0);
        assert!(math.height_emu > 0);
    }

    // ── Hard LaTeX construct tests ────────────────────────────────────

    /// Helper: assert a LaTeX expression renders successfully
    fn assert_renders(name: &str, latex: &str) {
        let result = render_latex_to_svg(latex, true, "10pt");
        assert!(result.is_ok(), "{} failed: {:?}", name, result.err());
        let math = result.unwrap();
        assert!(math.width_emu > 0 && math.height_emu > 0, "{} produced zero-size output", name);
    }

    #[test]
    fn test_piecewise_via_array() {
        assert_renders(
            "piecewise",
            r"f(x) = \left\{ \begin{array}{ll} x^2 & \text{if } x \geq 0 \\ -x & \text{if } x < 0 \end{array} \right.",
        );
    }

    #[test]
    fn test_aligned_environment() {
        assert_renders(
            "aligned",
            r"\begin{aligned} a &= b + c \\ d &= e + f \end{aligned}",
        );
    }

    #[test]
    fn test_substack() {
        assert_renders(
            "substack",
            r"\sum_{\substack{0 < i < m \\ 0 < j < n}} P(i,j)",
        );
    }

    #[test]
    fn test_overbrace_underbrace() {
        assert_renders("overbrace", r"\overbrace{a + b + c}^{3}");
        assert_renders("underbrace", r"\underbrace{x + y + z}_{3}");
    }

    #[test]
    fn test_continued_fraction() {
        // ReX doesn't have \cfrac, use nested \frac instead
        assert_renders(
            "continued-frac",
            r"\frac{1}{1+\frac{1}{1+\frac{1}{1+x}}}",
        );
    }

    #[test]
    fn test_multi_index_tensor() {
        assert_renders("tensor", r"R^{\mu}{}_{\nu\rho\sigma}");
    }

    #[test]
    fn test_floor_ceil() {
        assert_renders("floor-ceil", r"\lfloor x \rfloor + \lceil y \rceil");
    }

    #[test]
    fn test_binomial() {
        assert_renders("binom", r"\binom{n}{k} = \frac{n!}{k!(n-k)!}");
    }

    #[test]
    fn test_determinant_vmatrix() {
        assert_renders(
            "vmatrix",
            r"\begin{vmatrix} a & b \\ c & d \end{vmatrix} = ad - bc",
        );
    }

    #[test]
    fn test_text_command() {
        assert_renders("text", r"x = 1 \text{ if } y > 0");
    }

    #[test]
    fn test_font_variants() {
        assert_renders("mathbb", r"\mathbb{R}^n");
        assert_renders("mathbf", r"\mathbf{v} = (v_1, v_2, v_3)");
        assert_renders("mathcal", r"\mathcal{L}\{f(t)\} = F(s)");
    }

    #[test]
    fn test_nth_roots_various() {
        assert_renders("sqrt3", r"\sqrt[3]{27} = 3");
        assert_renders("sqrt4", r"\sqrt[4]{256} = 4");
        assert_renders("sqrt5", r"\sqrt[5]{32} = 2");
        assert_renders("sqrt-nested", r"\sqrt[3]{\sqrt{x^2 + 1}}");
    }

    #[test]
    fn test_dirac_bra_ket() {
        assert_renders("dirac", r"\langle \psi | \hat{H} | \phi \rangle");
    }

    #[test]
    fn test_stirling_approximation() {
        assert_renders(
            "stirling",
            r"n! \approx \sqrt{2\pi n} \left(\frac{n}{e}\right)^n",
        );
    }

    #[test]
    fn test_fourier_transform() {
        assert_renders(
            "fourier",
            r"\hat{f}(\xi) = \int_{-\infty}^{\infty} f(x) e^{-2\pi i x \xi} dx",
        );
    }

    #[test]
    fn test_laplacian() {
        assert_renders(
            "laplacian",
            r"\nabla^2 f = \frac{\partial^2 f}{\partial x^2} + \frac{\partial^2 f}{\partial y^2}",
        );
    }

    #[test]
    fn test_matrix_variants() {
        assert_renders("Bmatrix", r"\begin{Bmatrix} a & b \\ c & d \end{Bmatrix}");
        assert_renders("Vmatrix", r"\begin{Vmatrix} a & b \\ c & d \end{Vmatrix}");
    }

    #[test]
    fn test_multi_line_aligned() {
        assert_renders(
            "multi-aligned",
            r"\begin{aligned} f(x) &= x^2 + 2x + 1 \\ &= (x+1)^2 \end{aligned}",
        );
    }

    #[test]
    fn test_nested_fractions_deep() {
        assert_renders(
            "deep-nested",
            r"\frac{1}{1 + \frac{1}{1 + \frac{1}{1 + \frac{1}{x}}}}",
        );
    }

    #[test]
    fn test_sum_product_combined() {
        assert_renders(
            "sum-prod",
            r"\sum_{k=1}^{n} k^2 = \frac{n(n+1)(2n+1)}{6}",
        );
    }

    #[test]
    fn test_limit_multivar() {
        assert_renders(
            "limit-multivar",
            r"\lim_{(x,y) \to (0,0)} \frac{xy}{x^2 + y^2}",
        );
    }

    #[test]
    fn test_euler_product() {
        assert_renders(
            "euler-product",
            r"\zeta(s) = \sum_{n=1}^{\infty} \frac{1}{n^s} = \prod_{p} \frac{1}{1 - p^{-s}}",
        );
    }

    #[test]
    fn test_bayes_theorem() {
        assert_renders(
            "bayes",
            r"P(A|B) = \frac{P(B|A) \cdot P(A)}{P(B)}",
        );
    }

    #[test]
    fn test_maxwell_full() {
        assert_renders(
            "maxwell-curl",
            r"\nabla \times \vec{B} = \mu_0 \vec{J} + \mu_0 \epsilon_0 \frac{\partial \vec{E}}{\partial t}",
        );
    }

    #[test]
    fn test_taylor_series() {
        assert_renders(
            "taylor",
            r"f(x) = \sum_{n=0}^{\infty} \frac{f^{(n)}(a)}{n!}(x-a)^n",
        );
    }

    #[test]
    fn test_residue_theorem() {
        assert_renders(
            "residue",
            r"\oint_{\gamma} f(z) \, dz = 2\pi i \sum_{k=1}^{n} \text{Res}(f, a_k)",
        );
    }
}
