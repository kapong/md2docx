//! Mermaid diagram rendering using mermaid-rs-renderer
//!
//! Pure Rust implementation - no browser required.
//! 500-1000x faster than mermaid-cli.

pub mod config;
pub use config::MermaidConfig;

use crate::error::Error;

/// Render mermaid diagram to SVG string with text converted to paths
///
/// This ensures the SVG renders correctly in Microsoft Word, which has
/// limited font support for embedded SVGs. By converting text to paths,
/// the text becomes vector shapes that Word can always render.
///
/// # Arguments
/// * `content` - The mermaid diagram source code
///
/// # Returns
/// SVG string with text converted to paths
///
/// # Errors
/// Returns Error if rendering fails
pub fn render_to_svg(content: &str) -> Result<String, Error> {
    // First render to SVG using mermaid-rs-renderer
    let svg = mermaid_rs_renderer::render(content).map_err(|e| Error::Mermaid(e.to_string()))?;

    // Convert text to paths for Word compatibility
    // This requires usvg and fontdb (enabled via mermaid-png feature)
    #[cfg(feature = "mermaid-png")]
    {
        convert_text_to_paths(&svg)
    }
    #[cfg(not(feature = "mermaid-png"))]
    {
        Ok(svg)
    }
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
    // First get the SVG with text converted to paths
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

    // Parse SVG (text-to-path conversion already happened in render_to_svg)
    let opt = Options::default();
    let tree =
        Tree::from_str(svg, &opt).map_err(|e| Error::Mermaid(format!("SVG parse error: {}", e)))?;

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

    let opt = Options::default();
    let tree =
        Tree::from_str(svg, &opt).map_err(|e| Error::Mermaid(format!("SVG parse error: {}", e)))?;

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
        assert!(result.is_ok());
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
        assert!(svg.contains("</svg>"));
    }

    #[test]
    #[cfg(feature = "mermaid-png")]
    fn test_render_to_png() {
        let diagram = "flowchart LR; A-->B-->C";
        let result = render_to_png(diagram, 2.0);
        assert!(result.is_ok());
        let png = result.unwrap();
        // PNG magic bytes
        assert_eq!(&png[0..4], &[0x89, 0x50, 0x4E, 0x47]);
    }

    #[test]
    #[cfg(feature = "mermaid-png")]
    fn test_text_to_path_conversion() {
        let diagram = "flowchart LR; A[Hello World] --> B";
        let result = render_to_svg(diagram);
        assert!(result.is_ok());
        let svg = result.unwrap();
        // Should contain path elements and NOT contain "Hello World" as text if conversion worked
        assert!(svg.contains("<path"));
        assert!(!svg.contains(">Hello World<"));
    }
}
