//! Image template extraction from DOCX files
//!
//! Extracts image caption styling from a DOCX file.
//! The file should contain a sample image with a caption.

use crate::error::{Error, Result};
use std::path::Path;

/// Represents an extracted image template
#[derive(Debug, Clone)]
pub struct ImageTemplate {
    /// Caption style
    pub caption: ImageCaptionStyle,
    /// Image border style (if any)
    pub border: Option<ImageBorderStyle>,
    /// Default image width percentage (0-100)
    pub default_width_percent: u32,
    /// Image alignment: "left", "center", "right"
    pub alignment: String,
}

/// Image caption style
#[derive(Debug, Clone)]
pub struct ImageCaptionStyle {
    /// Caption position: "top" or "bottom" (images usually have bottom)
    pub position: String,
    /// Caption prefix (e.g., "Figure" or "รูปที่")
    pub prefix: String,
    /// Font family
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Font color (hex)
    pub font_color: String,
    /// Whether text is bold
    pub bold: bool,
    /// Whether text is italic
    pub italic: bool,
    /// Text alignment
    pub alignment: String,
    /// Spacing before caption in twips
    pub spacing_before: u32,
    /// Spacing after caption in twips
    pub spacing_after: u32,
}

/// Image border style
#[derive(Debug, Clone)]
pub struct ImageBorderStyle {
    /// Border style type: "single", "double", "none"
    pub style: String,
    /// Border color (hex)
    pub color: String,
    /// Border width in eighths of a point
    pub width: u32,
}

impl Default for ImageCaptionStyle {
    fn default() -> Self {
        Self {
            position: "bottom".to_string(),
            prefix: "Figure".to_string(),
            font_family: "Calibri".to_string(),
            font_size: 22, // 11pt
            font_color: "#4a5568".to_string(),
            bold: false,
            italic: true,
            alignment: "center".to_string(),
            spacing_before: 120, // 6pt
            spacing_after: 120,  // 6pt
        }
    }
}

impl Default for ImageTemplate {
    fn default() -> Self {
        Self {
            caption: ImageCaptionStyle::default(),
            border: None,
            default_width_percent: 80,
            alignment: "center".to_string(),
        }
    }
}

impl ImageTemplate {
    /// Create a new image template with default styles
    pub fn new() -> Self {
        Self::default()
    }

    /// Set caption prefix (e.g., "Figure" or "รูปที่")
    pub fn with_caption_prefix(mut self, prefix: impl Into<String>) -> Self {
        self.caption.prefix = prefix.into();
        self
    }

    /// Set caption position ("top" or "bottom")
    pub fn with_caption_position(mut self, position: impl Into<String>) -> Self {
        self.caption.position = position.into();
        self
    }

    /// Set default image width percentage
    pub fn with_default_width(mut self, percent: u32) -> Self {
        self.default_width_percent = percent.min(100);
        self
    }

    /// Format a caption with the given number and text
    ///
    /// # Arguments
    /// * `number` - The figure number (e.g., "1.2")
    /// * `text` - The caption text
    ///
    /// # Returns
    /// Formatted caption string (e.g., "Figure 1.2: Caption text")
    pub fn format_caption(&self, number: &str, text: &str) -> String {
        format!("{} {}: {}", self.caption.prefix, number, text)
    }
}

/// Extract image template from a DOCX file
///
/// This function reads a DOCX file and extracts image caption styling
/// from the sample image with caption.
///
/// # Arguments
/// * `path` - Path to the image.docx file
///
/// # Returns
/// The extracted `ImageTemplate`
///
/// # Example
/// ```rust,no_run
/// use md2docx::template::extract::extract_image;
/// use std::path::Path;
///
/// let image_template = extract_image(Path::new("my-template/image.docx")).unwrap();
/// println!("Caption prefix: {}", image_template.caption.prefix);
/// ```
pub fn extract(path: &Path) -> Result<ImageTemplate> {
    // TODO: Implement actual DOCX parsing
    // For now, return default template as a placeholder

    if !path.exists() {
        return Err(Error::Template(format!(
            "Image template file not found: {}",
            path.display()
        )));
    }

    // Placeholder implementation - will be replaced with actual extraction
    Ok(ImageTemplate::default())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_image_template_default() {
        let template = ImageTemplate::default();

        assert_eq!(template.caption.position, "bottom");
        assert_eq!(template.caption.prefix, "Figure");
        assert!(template.caption.italic);
        assert_eq!(template.caption.alignment, "center");
        assert_eq!(template.default_width_percent, 80);
        assert_eq!(template.alignment, "center");
    }

    #[test]
    fn test_image_template_builder() {
        let template = ImageTemplate::new()
            .with_caption_prefix("รูปที่")
            .with_caption_position("top")
            .with_default_width(100);

        assert_eq!(template.caption.prefix, "รูปที่");
        assert_eq!(template.caption.position, "top");
        assert_eq!(template.default_width_percent, 100);
    }

    #[test]
    fn test_format_caption() {
        let template = ImageTemplate::default();
        let caption = template.format_caption("1.2", "System Architecture");

        assert_eq!(caption, "Figure 1.2: System Architecture");
    }

    #[test]
    fn test_format_caption_thai() {
        let template = ImageTemplate::new().with_caption_prefix("รูปที่");
        let caption = template.format_caption("1.2", "โครงสร้างระบบ");

        assert_eq!(caption, "รูปที่ 1.2: โครงสร้างระบบ");
    }

    #[test]
    fn test_extract_file_not_found() {
        let result = extract(Path::new("/nonexistent/image.docx"));
        assert!(result.is_err());
    }
}
