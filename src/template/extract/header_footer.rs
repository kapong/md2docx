//! Header/Footer template extraction from DOCX files
//!
//! Extracts header and footer content with placeholders from a DOCX file.

use crate::error::{Error, Result};
use std::path::Path;

/// Represents an extracted header/footer template
#[derive(Debug, Clone)]
pub struct HeaderFooterTemplate {
    /// Header content
    pub header: HeaderContent,
    /// Footer content
    pub footer: FooterContent,
    /// Whether first page is different (no header/footer on cover)
    pub different_first_page: bool,
}

/// Header content with left, center, right sections
#[derive(Debug, Clone)]
pub struct HeaderContent {
    /// Left-aligned content (may contain placeholders)
    pub left: String,
    /// Center-aligned content (may contain placeholders)
    pub center: String,
    /// Right-aligned content (may contain placeholders)
    pub right: String,
    /// Font family
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Font color (hex)
    pub font_color: String,
}

/// Footer content with left, center, right sections
#[derive(Debug, Clone)]
pub struct FooterContent {
    /// Left-aligned content (may contain placeholders)
    pub left: String,
    /// Center-aligned content (may contain placeholders)
    pub center: String,
    /// Right-aligned content (may contain placeholders)
    pub right: String,
    /// Font family
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Font color (hex)
    pub font_color: String,
}

impl Default for HeaderContent {
    fn default() -> Self {
        Self {
            left: "{{title}}".to_string(),
            center: "".to_string(),
            right: "{{chapter}}".to_string(),
            font_family: "Calibri".to_string(),
            font_size: 20, // 10pt
            font_color: "#4a5568".to_string(),
        }
    }
}

impl Default for FooterContent {
    fn default() -> Self {
        Self {
            left: "".to_string(),
            center: "{{page}}".to_string(),
            right: "".to_string(),
            font_family: "Calibri".to_string(),
            font_size: 20, // 10pt
            font_color: "#4a5568".to_string(),
        }
    }
}

impl Default for HeaderFooterTemplate {
    fn default() -> Self {
        Self {
            header: HeaderContent::default(),
            footer: FooterContent::default(),
            different_first_page: true,
        }
    }
}

impl HeaderFooterTemplate {
    /// Create a new header/footer template with default content
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if header has any placeholders
    pub fn header_has_placeholders(&self) -> bool {
        self.header.left.contains("{{")
            || self.header.center.contains("{{")
            || self.header.right.contains("{{")
    }

    /// Check if footer has any placeholders
    pub fn footer_has_placeholders(&self) -> bool {
        self.footer.left.contains("{{")
            || self.footer.center.contains("{{")
            || self.footer.right.contains("{{")
    }

    /// Get all unique placeholder keys from header and footer
    pub fn all_placeholders(&self) -> Vec<String> {
        use crate::template::placeholder::extract_placeholders;
        use std::collections::HashSet;

        let mut keys = HashSet::new();

        // Extract from header
        for key in extract_placeholders(&self.header.left) {
            keys.insert(key);
        }
        for key in extract_placeholders(&self.header.center) {
            keys.insert(key);
        }
        for key in extract_placeholders(&self.header.right) {
            keys.insert(key);
        }

        // Extract from footer
        for key in extract_placeholders(&self.footer.left) {
            keys.insert(key);
        }
        for key in extract_placeholders(&self.footer.center) {
            keys.insert(key);
        }
        for key in extract_placeholders(&self.footer.right) {
            keys.insert(key);
        }

        keys.into_iter().collect()
    }
}

/// Extract header/footer template from a DOCX file
///
/// This function reads a DOCX file and extracts header and footer content
/// with their formatting.
///
/// # Arguments
/// * `path` - Path to the header-footer.docx file
///
/// # Returns
/// The extracted `HeaderFooterTemplate`
///
/// # Example
/// ```rust,no_run
/// use md2docx::template::extract::extract_header_footer;
/// use std::path::Path;
///
/// let hf = extract_header_footer(Path::new("my-template/header-footer.docx")).unwrap();
/// println!("Header left: {}", hf.header.left);
/// println!("Footer center: {}", hf.footer.center);
/// ```
pub fn extract(path: &Path) -> Result<HeaderFooterTemplate> {
    // TODO: Implement actual DOCX parsing
    // For now, return default template as a placeholder

    if !path.exists() {
        return Err(Error::Template(format!(
            "Header/footer template file not found: {}",
            path.display()
        )));
    }

    // Placeholder implementation - will be replaced with actual extraction
    Ok(HeaderFooterTemplate::default())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_footer_template_default() {
        let template = HeaderFooterTemplate::default();

        assert_eq!(template.header.left, "{{title}}");
        assert_eq!(template.header.right, "{{chapter}}");
        assert_eq!(template.footer.center, "{{page}}");
        assert!(template.different_first_page);
    }

    #[test]
    fn test_header_has_placeholders() {
        let template = HeaderFooterTemplate::default();
        assert!(template.header_has_placeholders());

        let template_no_placeholders = HeaderFooterTemplate {
            header: HeaderContent {
                left: "Static Text".to_string(),
                right: "More Static Text".to_string(),
                ..Default::default()
            },
            ..Default::default()
        };
        assert!(!template_no_placeholders.header_has_placeholders());
    }

    #[test]
    fn test_footer_has_placeholders() {
        let template = HeaderFooterTemplate::default();
        assert!(template.footer_has_placeholders());

        let template_no_placeholders = HeaderFooterTemplate {
            footer: FooterContent {
                center: "Static".to_string(),
                ..Default::default()
            },
            ..Default::default()
        };
        assert!(!template_no_placeholders.footer_has_placeholders());
    }

    #[test]
    fn test_all_placeholders() {
        let template = HeaderFooterTemplate::default();
        let placeholders = template.all_placeholders();

        assert!(placeholders.contains(&"title".to_string()));
        assert!(placeholders.contains(&"chapter".to_string()));
        assert!(placeholders.contains(&"page".to_string()));
    }

    #[test]
    fn test_extract_file_not_found() {
        let result = extract(Path::new("/nonexistent/header-footer.docx"));
        assert!(result.is_err());
    }
}
