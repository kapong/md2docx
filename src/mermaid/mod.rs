//! Mermaid diagram rendering using mermaid-rs-renderer
//!
//! Pure Rust implementation - no browser required.
//! 500-1000x faster than mermaid-cli.

pub mod config;
pub use config::MermaidConfig;

use crate::error::Error;

/// Render mermaid diagram to PNG bytes
///
/// # Arguments
/// * `content` - The mermaid diagram source code
///
/// # Returns
/// PNG image bytes
///
/// # Errors
/// Returns Error if rendering fails
pub fn render_to_png(_content: &str) -> Result<Vec<u8>, Error> {
    // mermaid-rs-renderer with default features includes PNG support via resvg
    // But we disabled default-features, so we need to handle conversion if we want PNG.
    // For now, let's just use SVG since DOCX supports it.
    // Actually, mermaid-rs-renderer renders to SVG.

    // Fallback or implementation for PNG if needed later
    Err(Error::NotImplemented(
        "PNG rendering for Mermaid not yet implemented".to_string(),
    ))
}

/// Render mermaid diagram to SVG bytes
pub fn render_to_svg(content: &str) -> Result<String, Error> {
    mermaid_rs_renderer::render(content).map_err(|e| Error::Mermaid(e.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_render_simple_flowchart() {
        let diagram = "flowchart LR; A-->B-->C";
        let result = render_to_svg(diagram);
        // Note: this test might fail if the library isn't actually loaded or has issues
        // But for now it's according to instructions
        assert!(result.is_ok());
        let svg = result.unwrap();
        assert!(svg.contains("<svg"));
        assert!(svg.contains("</svg>"));
    }
}
