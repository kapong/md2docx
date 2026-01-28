//! Table template extraction from DOCX files
//!
//! Extracts table styling from a sample table in a DOCX file.
//! The sample table should have:
//! - Row 1: Header row style
//! - Row 2: Odd row style
//! - Row 3: Even row style
//! - Row 4+: First column style example

use crate::error::{Error, Result};
use std::path::Path;

/// Represents an extracted table template
#[derive(Debug, Clone)]
pub struct TableTemplate {
    /// Header row style
    pub header: RowStyle,
    /// Odd row style (row 1, 3, 5, ...)
    pub row_odd: RowStyle,
    /// Even row style (row 2, 4, 6, ...)
    pub row_even: RowStyle,
    /// First column cell style
    pub first_column: CellStyle,
    /// Other columns cell style
    pub other_columns: CellStyle,
    /// Border styles
    pub borders: BorderStyles,
    /// Caption style
    pub caption: TableCaptionStyle,
}

/// Row style properties
#[derive(Debug, Clone)]
pub struct RowStyle {
    /// Background color (hex, e.g., "#1a365d")
    pub background_color: Option<String>,
    /// Font family name
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Font color (hex)
    pub font_color: String,
    /// Whether text is bold
    pub bold: bool,
    /// Whether text is italic
    pub italic: bool,
}

impl Default for RowStyle {
    fn default() -> Self {
        Self {
            background_color: None,
            font_family: "Calibri".to_string(),
            font_size: 22, // 11pt
            font_color: "#000000".to_string(),
            bold: false,
            italic: false,
        }
    }
}

/// Cell style properties
#[derive(Debug, Clone)]
pub struct CellStyle {
    /// Font family name
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Font color (hex)
    pub font_color: String,
    /// Whether text is bold
    pub bold: bool,
    /// Whether text is italic
    pub italic: bool,
    /// Text alignment: "left", "center", "right"
    pub alignment: String,
    /// Vertical alignment: "top", "center", "bottom"
    pub vertical_alignment: String,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self {
            font_family: "Calibri".to_string(),
            font_size: 22, // 11pt
            font_color: "#000000".to_string(),
            bold: false,
            italic: false,
            alignment: "left".to_string(),
            vertical_alignment: "center".to_string(),
        }
    }
}

/// Border styles
#[derive(Debug, Clone)]
pub struct BorderStyles {
    /// Top border style
    pub top: BorderStyle,
    /// Bottom border style
    pub bottom: BorderStyle,
    /// Left border style
    pub left: BorderStyle,
    /// Right border style
    pub right: BorderStyle,
    /// Inside horizontal borders
    pub inside_h: BorderStyle,
    /// Inside vertical borders
    pub inside_v: BorderStyle,
}

/// Individual border style
#[derive(Debug, Clone)]
pub struct BorderStyle {
    /// Border style type: "single", "double", "dashed", "none"
    pub style: String,
    /// Border color (hex)
    pub color: String,
    /// Border width in eighths of a point
    pub width: u32,
}

impl Default for BorderStyle {
    fn default() -> Self {
        Self {
            style: "single".to_string(),
            color: "#000000".to_string(),
            width: 4, // 0.5pt
        }
    }
}

impl Default for BorderStyles {
    fn default() -> Self {
        Self {
            top: BorderStyle::default(),
            bottom: BorderStyle::default(),
            left: BorderStyle::default(),
            right: BorderStyle::default(),
            inside_h: BorderStyle::default(),
            inside_v: BorderStyle::default(),
        }
    }
}

/// Table caption style
#[derive(Debug, Clone)]
pub struct TableCaptionStyle {
    /// Caption position: "top" or "bottom" (tables usually have top)
    pub position: String,
    /// Caption prefix (e.g., "Table" or "ตารางที่")
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

impl Default for TableCaptionStyle {
    fn default() -> Self {
        Self {
            position: "top".to_string(),
            prefix: "Table".to_string(),
            font_family: "Calibri".to_string(),
            font_size: 22, // 11pt
            font_color: "#4a5568".to_string(),
            bold: false,
            italic: true,
            alignment: "left".to_string(),
            spacing_before: 120, // 6pt
            spacing_after: 120,  // 6pt
        }
    }
}

impl Default for TableTemplate {
    fn default() -> Self {
        Self {
            header: RowStyle {
                background_color: Some("#1a365d".to_string()),
                font_family: "Calibri".to_string(),
                font_size: 22,
                font_color: "#FFFFFF".to_string(),
                bold: true,
                italic: false,
            },
            row_odd: RowStyle::default(),
            row_even: RowStyle {
                background_color: Some("#f7fafc".to_string()),
                ..Default::default()
            },
            first_column: CellStyle {
                bold: true,
                ..Default::default()
            },
            other_columns: CellStyle::default(),
            borders: BorderStyles::default(),
            caption: TableCaptionStyle::default(),
        }
    }
}

impl TableTemplate {
    /// Create a new table template with default styles
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the appropriate row style based on row index (0-based)
    ///
    /// - Index 0: Header style
    /// - Odd indices (1, 3, 5...): Odd row style
    /// - Even indices (2, 4, 6...): Even row style
    pub fn row_style_for_index(&self, index: usize) -> &RowStyle {
        match index {
            0 => &self.header,
            i if i % 2 == 1 => &self.row_odd,
            _ => &self.row_even,
        }
    }

    /// Get the appropriate cell style based on column index
    ///
    /// - Index 0: First column style
    /// - Others: Other columns style
    pub fn cell_style_for_column(&self, index: usize) -> &CellStyle {
        if index == 0 {
            &self.first_column
        } else {
            &self.other_columns
        }
    }
}

/// Extract table template from a DOCX file
///
/// This function reads a DOCX file and extracts table styling from
/// the sample table. The table should have at least 4 rows:
/// - Row 1: Header row style
/// - Row 2: Odd row style
/// - Row 3: Even row style
/// - Row 4: First column style example
///
/// # Arguments
/// * `path` - Path to the table.docx file
///
/// # Returns
/// The extracted `TableTemplate`
///
/// # Example
/// ```rust,no_run
/// use md2docx::template::extract::extract_table;
/// use std::path::Path;
///
/// let table_template = extract_table(Path::new("my-template/table.docx")).unwrap();
/// println!("Header font: {}", table_template.header.font_family);
/// ```
pub fn extract(path: &Path) -> Result<TableTemplate> {
    // TODO: Implement actual DOCX parsing
    // For now, return default template as a placeholder

    if !path.exists() {
        return Err(Error::Template(format!(
            "Table template file not found: {}",
            path.display()
        )));
    }

    // Placeholder implementation - will be replaced with actual extraction
    Ok(TableTemplate::default())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_template_default() {
        let template = TableTemplate::default();

        // Header should have dark background and white text
        assert_eq!(
            template.header.background_color,
            Some("#1a365d".to_string())
        );
        assert_eq!(template.header.font_color, "#FFFFFF");
        assert!(template.header.bold);

        // Even rows should have light gray background
        assert_eq!(
            template.row_even.background_color,
            Some("#f7fafc".to_string())
        );

        // First column should be bold
        assert!(template.first_column.bold);
    }

    #[test]
    fn test_row_style_for_index() {
        let template = TableTemplate::default();

        assert_eq!(template.row_style_for_index(0).font_color, "#FFFFFF"); // Header
        assert_eq!(template.row_style_for_index(1).font_color, "#000000"); // Odd
        assert_eq!(
            template.row_style_for_index(2).background_color,
            Some("#f7fafc".to_string())
        ); // Even
        assert_eq!(template.row_style_for_index(3).font_color, "#000000"); // Odd
    }

    #[test]
    fn test_cell_style_for_column() {
        let template = TableTemplate::default();

        assert!(template.cell_style_for_column(0).bold); // First column
        assert!(!template.cell_style_for_column(1).bold); // Other columns
        assert!(!template.cell_style_for_column(2).bold); // Other columns
    }

    #[test]
    fn test_caption_style_default() {
        let caption = TableCaptionStyle::default();

        assert_eq!(caption.position, "top");
        assert_eq!(caption.prefix, "Table");
        assert!(caption.italic);
    }

    #[test]
    fn test_extract_file_not_found() {
        let result = extract(Path::new("/nonexistent/table.docx"));
        assert!(result.is_err());
    }
}
