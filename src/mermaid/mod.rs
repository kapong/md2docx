//! Mermaid diagram rendering using mermaid-rs-renderer
//!
//! Pure Rust implementation - no browser required.
//! 500-1400x faster than mermaid-cli.
//!
//! v0.2.0 supports all 23 diagram types natively: flowchart, sequence, class,
//! state, ER, pie, gantt, journey, timeline, mindmap, gitGraph, xychart,
//! quadrant, sankey, kanban, C4, block, architecture, requirement, zenuml,
//! packet, radar, treemap.

pub mod config;
pub use config::MermaidConfig;

use crate::error::Error;
use once_cell::sync::Lazy;
use regex::Regex;

/// Padding factor for SVG canvas (1.0 = no extra padding)
const SVG_PADDING_FACTOR: f64 = 1.0;

use std::panic;

/// Static regex for pipe-separated edge labels: -->|label| or --|label|->, etc.
static PIPE_LABEL_RE: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r#"(-{1,2}?=?>?|->?|\.-\.)\|[^|]+\|"#).expect("pipe label regex is valid")
});

/// Static regex for bracket edge labels: -->[label] or --> [label]
static BRACKET_LABEL_RE: Lazy<Regex> = Lazy::new(|| {
    Regex::new(r#"(-{1,2}?=?>?|->?)\s*\[([^\]]+)\]"#).expect("bracket label regex is valid")
});

/// Static regex for extracting SVG width attribute
static WIDTH_RE: Lazy<Regex> =
    Lazy::new(|| Regex::new(r#"width="([^"]+)""#).expect("width regex is valid"));

/// Static regex for extracting SVG height attribute
static HEIGHT_RE: Lazy<Regex> =
    Lazy::new(|| Regex::new(r#"height="([^"]+)""#).expect("height regex is valid"));

/// Render mermaid diagram to SVG string with text converted to paths
///
/// This ensures the SVG renders correctly in Microsoft Word, which has
/// limited font support for embedded SVGs. By converting text to paths,
/// the text becomes vector shapes that Word can always render.
///
/// If the diagram contains unsupported features (like edge labels with Thai text),
/// it will attempt to simplify the diagram and render a basic version.
///
/// # Arguments
/// * `content` - The mermaid diagram source code
///
/// # Returns
/// SVG string with text converted to paths and proper canvas sizing
///
/// # Errors
/// Returns Error if rendering fails even after fallback
pub fn render_to_svg(content: &str) -> Result<String, Error> {
    // v0.2.0 of mermaid-rs-renderer supports all 23 diagram types natively:
    // flowchart, sequence, class, state, ER, pie, gantt, journey, timeline,
    // mindmap, gitGraph, xychart, quadrant, sankey, kanban, C4, block,
    // architecture, requirement, zenuml, packet, radar, treemap.

    // Try normal rendering first
    match try_render_to_svg(content) {
        Ok(svg) => Ok(svg),
        Err(e) => {
            // If normal rendering fails, try stripping edge labels
            let simplified = strip_edge_labels(content);
            if simplified != content {
                eprintln!("Warning: Mermaid diagram contains unsupported features (edge labels). Rendering simplified version without labels.");
                try_render_to_svg(&simplified)
            } else {
                Err(e)
            }
        }
    }
}

/// Sanitize SVG output from mermaid-rs-renderer for usvg compatibility.
///
/// mermaid-rs-renderer v0.2.0 may produce `font-family` attributes with unescaped
/// inner double quotes (e.g., `font-family="..., "Segoe UI", sans-serif"`) which
/// are invalid XML and cause usvg to fail parsing. This function finds such
/// inner quotes and replaces them with single quotes so the attribute becomes
/// valid XML: `font-family="..., 'Segoe UI', ..."`.
fn sanitize_svg_for_usvg(svg: &str) -> String {
    let marker = "font-family=\"";
    let mut result = String::with_capacity(svg.len());
    let mut remaining = svg;

    while let Some(marker_pos) = remaining.find(marker) {
        // Copy everything up to and including the opening quote
        result.push_str(&remaining[..marker_pos + marker.len()]);
        remaining = &remaining[marker_pos + marker.len()..];

        // Scan the attribute value, replacing inner quotes with single quotes
        loop {
            if let Some(quote_pos) = remaining.find('"') {
                // Copy everything before this quote
                result.push_str(&remaining[..quote_pos]);

                // Is this the real closing quote? In valid XML, the closing quote
                // is followed by whitespace, '>', '/', or end of string.
                let after = &remaining[quote_pos + 1..];
                let is_closing = after.is_empty()
                    || after.starts_with(' ')
                    || after.starts_with('>')
                    || after.starts_with('/')
                    || after.starts_with('\n')
                    || after.starts_with('\r')
                    || after.starts_with('\t');

                if is_closing {
                    result.push('"');
                    remaining = &remaining[quote_pos + 1..];
                    break;
                } else {
                    // Inner quote (e.g., the " before "Segoe UI") → single quote
                    result.push('\'');
                    remaining = &remaining[quote_pos + 1..];
                }
            } else {
                // No more quotes — malformed, just copy the rest
                result.push_str(remaining);
                remaining = "";
                break;
            }
        }
    }

    result.push_str(remaining);
    result
}

/// Attempt to render mermaid diagram to SVG
///
/// Wraps the actual renderer in catch_unwind to prevent panics.
fn try_render_to_svg(content: &str) -> Result<String, Error> {
    // Wrap in catch_unwind to prevent panics in the renderer from crashing the tool
    let content_owned = content.to_string();
    let result = panic::catch_unwind(move || mermaid_rs_renderer::render(&content_owned));

    let svg = match result {
        Ok(Ok(svg)) => svg,
        Ok(Err(e)) => return Err(Error::Mermaid(e.to_string())),
        Err(_) => {
            return Err(Error::Mermaid(
                "Mermaid renderer panicked (likely due to Unicode or syntax issues)".to_string(),
            ))
        }
    };

    // Sanitize SVG: fix escaped quotes in font-family attributes for usvg compatibility
    let svg = sanitize_svg_for_usvg(&svg);

    // Convert text to paths for Word compatibility (SVG output only).
    // When used for PNG conversion, svg_to_png handles text rendering via usvg fontdb.
    #[cfg(feature = "mermaid-png")]
    let svg = match convert_text_to_paths(&svg) {
        Ok(converted) => converted,
        Err(_) => svg, // Keep original SVG if conversion fails
    };

    // Add padding to canvas to prevent arrow/edge clipping
    let svg = add_canvas_padding(&svg)?;

    Ok(svg)
}

/// Strip edge labels from mermaid diagram for compatibility
///
/// Converts edge labels like `A -->|label| B` to simple arrows `A --> B`.
/// This is a fallback when the mermaid renderer doesn't support certain features
/// like edge labels with Unicode/Thai text.
///
/// # Patterns handled
/// - `-->|label|` - solid arrow with label
/// - `--|label|->` - dashed arrow with label
/// - `==>|label|` - thick arrow with label
/// - `-->[label]` - bracketed label (with or without space)
///
/// # Returns
/// Simplified mermaid code without edge labels
fn strip_edge_labels(content: &str) -> String {
    // Replace pipe labels with simple arrow
    let result = PIPE_LABEL_RE.replace_all(content, "$1");

    // Replace bracket labels with simple arrow
    let result = BRACKET_LABEL_RE.replace_all(&result, "$1");

    result.to_string()
}

/// Convert SVG text elements to path elements using usvg
///
/// This ensures the SVG renders identically everywhere regardless of font availability.
#[cfg(feature = "mermaid-png")]
fn convert_text_to_paths(svg: &str) -> Result<String, Error> {
    use usvg::{fontdb, Options, Tree, WriteOptions};

    // Create options with font database
    let mut opt = Options::default();

    // Load system fonts for text-to-path conversion
    let mut font_db = fontdb::Database::new();
    font_db.load_system_fonts();
    opt.fontdb = std::sync::Arc::new(font_db);

    // Parse SVG
    let tree =
        Tree::from_str(svg, &opt).map_err(|e| Error::Mermaid(format!("SVG parse error: {}", e)))?;

    // Write back to SVG string - usvg automatically converts text to paths
    // in its internal tree representation during parsing if fontdb is provided.
    Ok(tree.to_string(&WriteOptions::default()))
}

/// Add padding to SVG canvas to prevent clipping of arrows and edges
///
/// This increases the width and height attributes while keeping the viewBox,
/// effectively adding margin around the diagram content.
fn add_canvas_padding(svg: &str) -> Result<String, Error> {
    // Extract current dimensions
    let width_caps = WIDTH_RE
        .captures(svg)
        .ok_or_else(|| Error::Mermaid("No width attribute found".to_string()))?;
    let height_caps = HEIGHT_RE
        .captures(svg)
        .ok_or_else(|| Error::Mermaid("No height attribute found".to_string()))?;

    let width_str = width_caps
        .get(1)
        .expect("width regex capture group 1 must exist")
        .as_str();
    let height_str = height_caps
        .get(1)
        .expect("height regex capture group 1 must exist")
        .as_str();

    // Parse dimensions (handle units like "px", "pt", or unitless)
    let width: f64 = parse_dimension(width_str)?;
    let height: f64 = parse_dimension(height_str)?;

    // Calculate new dimensions with padding
    let new_width = width * SVG_PADDING_FACTOR;
    let new_height = height * SVG_PADDING_FACTOR;

    // Replace width and height attributes
    let new_width_attr = format!(r#"width="{}""#, format_dimension(new_width, width_str));
    let new_height_attr = format!(r#"height="{}""#, format_dimension(new_height, height_str));

    let result = WIDTH_RE.replace(svg, new_width_attr.as_str());
    let result = HEIGHT_RE.replace(&result, new_height_attr.as_str());

    Ok(result.to_string())
}

/// Parse a dimension value (e.g., "100px", "100", "100.5")
fn parse_dimension(s: &str) -> Result<f64, Error> {
    let num_str: String = s
        .chars()
        .take_while(|c| c.is_ascii_digit() || *c == '.' || *c == '-')
        .collect();

    num_str
        .parse::<f64>()
        .map_err(|_| Error::Mermaid(format!("Invalid dimension: {}", s)))
}

/// Format a dimension value, preserving original units
fn format_dimension(value: f64, original: &str) -> String {
    // Extract unit from original
    let unit: String = original
        .chars()
        .skip_while(|c| c.is_ascii_digit() || *c == '.' || *c == '-')
        .collect();

    if unit.is_empty() {
        format!("{:.2}", value)
    } else {
        format!("{:.2}{}", value, unit)
    }
}

/// Render mermaid diagram to PNG bytes
///
/// This converts SVG to PNG for better compatibility with Microsoft Word versions
/// that don't support SVG, or when a raster image is preferred.
///
/// # Arguments
/// * `content` - The mermaid diagram source code
/// * `scale` - Scale factor for resolution (2.0 = 2x for high DPI)
///
/// # Returns
/// PNG image bytes
///
/// # Errors
/// Returns Error if rendering fails
#[cfg(feature = "mermaid-png")]
pub fn render_to_png(content: &str, scale: f32) -> Result<Vec<u8>, Error> {
    // First get the SVG with padding and text converted to paths
    let svg = render_to_svg(content)?;

    // Convert SVG to PNG
    svg_to_png(&svg, scale)
}

/// Render mermaid diagram to PNG bytes (without mermaid-png feature)
#[cfg(not(feature = "mermaid-png"))]
pub fn render_to_png(_content: &str, _scale: f32) -> Result<Vec<u8>, Error> {
    Err(Error::Mermaid(
        "PNG rendering requires 'mermaid-png' feature".to_string(),
    ))
}

/// Convert SVG string to PNG bytes using resvg
#[cfg(feature = "mermaid-png")]
fn svg_to_png(svg: &str, scale: f32) -> Result<Vec<u8>, Error> {
    use resvg::render;
    use usvg::{Options, Tree};

    // Sanitize SVG for usvg compatibility (safety net in case raw SVG is passed)
    let svg = sanitize_svg_for_usvg(svg);

    // Parse SVG (text-to-path conversion and padding already happened)
    let opt = Options::default();
    let tree = Tree::from_str(&svg, &opt)
        .map_err(|e| Error::Mermaid(format!("SVG parse error: {}", e)))?;

    // Get dimensions
    let size = tree.size();
    let width = (size.width() * scale) as u32;
    let height = (size.height() * scale) as u32;

    // Create pixmap
    let mut pixmap = tiny_skia::Pixmap::new(width, height)
        .ok_or_else(|| Error::Mermaid("Failed to create pixmap".to_string()))?;

    // Render SVG to pixmap
    render(
        &tree,
        tiny_skia::Transform::from_scale(scale, scale),
        &mut pixmap.as_mut(),
    );

    // Encode as PNG
    pixmap
        .encode_png()
        .map_err(|e| Error::Mermaid(format!("PNG encode error: {}", e)))
}

/// Get SVG dimensions
#[cfg(feature = "mermaid-png")]
pub fn get_svg_dimensions(svg: &str) -> Result<(u32, u32), Error> {
    use usvg::{Options, Tree};

    // Sanitize SVG for usvg compatibility
    let svg = sanitize_svg_for_usvg(svg);

    let opt = Options::default();
    let tree = Tree::from_str(&svg, &opt)
        .map_err(|e| Error::Mermaid(format!("SVG parse error: {}", e)))?;

    let size = tree.size();
    Ok((size.width() as u32, size.height() as u32))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_render_simple_flowchart() {
        let diagram = "flowchart LR; A-->B-->C";
        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "render_to_svg failed: {:?}", result.err());
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
        assert!(svg.contains("</svg>"));
    }

    #[test]
    #[cfg(feature = "mermaid-png")]
    fn test_render_to_png() {
        let diagram = "flowchart LR; A-->B-->C";
        let result = render_to_png(diagram, 2.0);
        assert!(result.is_ok(), "render_to_png failed: {:?}", result.err());
        let png = result.unwrap();
        // PNG magic bytes
        assert_eq!(&png[0..4], &[0x89, 0x50, 0x4E, 0x47]);
    }

    #[test]
    #[cfg(feature = "mermaid-png")]
    fn test_text_to_path_conversion() {
        let diagram = "flowchart LR; A[Hello World] --> B";
        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "render_to_svg failed: {:?}", result.err());
        let svg = result.unwrap();
        // After text-to-path conversion via usvg, text should be converted
        // to paths. The SVG should be valid and contain path elements.
        assert!(svg.contains("<svg"));
        assert!(svg.contains("<path"));
    }

    #[test]
    fn test_strip_edge_labels_pipe_format() {
        let input = "A -->|label| B";
        let output = strip_edge_labels(input);
        assert_eq!(output, "A --> B");
    }

    #[test]
    fn test_strip_edge_labels_multiple_pipe_labels() {
        let input = "A -->|x| B -->|y| C";
        let output = strip_edge_labels(input);
        assert_eq!(output, "A --> B --> C");
    }

    #[test]
    fn test_strip_edge_labels_bracket_format() {
        let input = "A --> [label] B";
        let output = strip_edge_labels(input);
        assert_eq!(output, "A --> B");
    }

    #[test]
    fn test_strip_edge_labels_bracket_format_no_space() {
        let input = "A -->[label] B";
        let output = strip_edge_labels(input);
        assert_eq!(output, "A --> B");
    }

    #[test]
    fn test_strip_edge_labels_mixed_formats() {
        let input = "A -->|pipe| B --> [bracket] C";
        let output = strip_edge_labels(input);
        assert_eq!(output, "A --> B --> C");
    }

    #[test]
    fn test_strip_edge_labels_complex_diagram() {
        let input = "flowchart TB
    subgraph Input
        MD[Markdown]
    end
    subgraph Output
        DOCX[DOCX]
    end
    MD -->|Process| DOCX";

        let output = strip_edge_labels(input);
        assert!(output.contains("MD --> DOCX"));
        assert!(!output.contains("|Process|"));
    }

    #[test]
    fn test_render_with_edge_labels_english() {
        // This diagram has edge labels but should render after stripping them
        let diagram = "flowchart LR
    Start -->|Pass| Success
    Start -->|Fail| Error";

        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "Should render after stripping edge labels");
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
    }

    #[test]
    fn test_render_thai_text() {
        // v0.2.0 supports Thai/Unicode text via fontdb
        let diagram = "flowchart LR
    A[เริ่มต้น] --> B[จบ]";

        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "Thai text should render in v0.2.0: {:?}", result.err());
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
    }

    #[test]
    fn test_sequence_diagram() {
        // v0.2.0 supports sequence diagrams natively
        let diagram = "sequenceDiagram\n    A->>B: Hello";
        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "sequenceDiagram should render in v0.2.0: {:?}", result.err());
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
    }

    #[test]
    fn test_er_diagram() {
        // v0.2.0 supports ER diagrams natively
        let diagram = "erDiagram\n    CUSTOMER ||--o{ ORDER : places";
        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "erDiagram should render in v0.2.0: {:?}", result.err());
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
    }

    #[test]
    fn test_pie_chart() {
        let diagram = "pie\n    title Languages\n    \"Rust\" : 45\n    \"Go\" : 30\n    \"Python\" : 25";
        let result = render_to_svg(diagram);
        assert!(result.is_ok(), "pie chart should render in v0.2.0: {:?}", result.err());
    }

    #[test]
    fn test_sanitize_svg_for_usvg() {
        // Simulate SVG with inner double quotes in font-family (as v0.2.0 produces)
        let input = "font-family=\"Georgia, -apple-system, \"Segoe UI\", sans-serif\" font-size=\"14\"";
        let sanitized = sanitize_svg_for_usvg(input);
        assert!(
            sanitized.contains("'Segoe UI'"),
            "Should have single-quoted font name, got: {}",
            sanitized
        );
        assert!(
            sanitized.contains("font-size=\"14\""),
            "Other attributes should be preserved"
        );
    }

    #[test]
    fn test_sanitize_svg_no_inner_quotes() {
        // SVG without the issue should pass through unchanged
        let input = "font-family=\"sans-serif\" font-size=\"14\"";
        let sanitized = sanitize_svg_for_usvg(input);
        assert_eq!(sanitized, input);
    }

    #[test]
    #[cfg(feature = "mermaid-png")]
    fn test_render_sequence_to_png() {
        // v0.2.0: all diagram types should render to PNG
        let diagram = "sequenceDiagram\n    Alice->>Bob: Hello";
        let result = render_to_png(diagram, 2.0);
        assert!(result.is_ok(), "sequenceDiagram PNG failed: {:?}", result.err());
        let png = result.unwrap();
        assert_eq!(&png[0..4], &[0x89, 0x50, 0x4E, 0x47]);
    }
}
