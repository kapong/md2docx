//! Table template extraction from DOCX files
//!
//! Extracts table styling from a sample table in a DOCX file.
//! The sample table should have:
//! - Row 1: Header row style
//! - Row 2: Odd row style
//! - Row 3: Even row style
//! - Row 4+: First column style example

use crate::error::{Error, Result};
use std::io::Read;
use std::path::Path;
use zip::ZipArchive;

/// Cell margins/padding in twips (1/20th of a point)
#[derive(Debug, Clone)]
pub struct CellMargins {
    /// Top margin in twips
    pub top: u32,
    /// Bottom margin in twips
    pub bottom: u32,
    /// Left margin in twips
    pub left: u32,
    /// Right margin in twips
    pub right: u32,
}

impl Default for CellMargins {
    fn default() -> Self {
        Self {
            top: 0,
            bottom: 0,
            left: 108, // Default Word value (~5.4pt)
            right: 108,
        }
    }
}

/// Paragraph spacing for table cells
#[derive(Debug, Clone)]
pub struct CellSpacing {
    /// Line height in twips (240 = single line)
    pub line: u32,
    /// Line rule: "auto", "exact", "atLeast"
    pub line_rule: String,
    /// Spacing before paragraph in twips
    pub before: u32,
    /// Spacing after paragraph in twips
    pub after: u32,
}

impl Default for CellSpacing {
    fn default() -> Self {
        Self {
            line: 240, // Single spacing
            line_rule: "auto".to_string(),
            before: 0,
            after: 0,
        }
    }
}

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
    /// Cell margins (padding)
    pub cell_margins: CellMargins,
    /// Cell paragraph spacing
    pub cell_spacing: CellSpacing,
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
#[derive(Debug, Clone, Default)]
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
            // Note: "Table" is the English default. When Language::Thai is used
            // and prefix is "Table", it will be replaced with "ตารางที่"
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
            cell_margins: CellMargins::default(),
            cell_spacing: CellSpacing::default(),
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
    if !path.exists() {
        return Err(Error::Template(format!(
            "Table template file not found: {}",
            path.display()
        )));
    }

    // Open DOCX as ZIP
    let file = std::fs::File::open(path)
        .map_err(|e| Error::Template(format!("Failed to open table template: {}", e)))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|e| Error::Template(format!("Failed to read table template as ZIP: {}", e)))?;

    // Read word/document.xml
    let mut document_xml = String::new();
    {
        let mut doc_file = archive
            .by_name("word/document.xml")
            .map_err(|e| Error::Template(format!("Failed to find document.xml: {}", e)))?;
        doc_file
            .read_to_string(&mut document_xml)
            .map_err(|e| Error::Template(format!("Failed to read document.xml: {}", e)))?;
    }

    // Parse XML and extract table styles
    extract_from_xml(&document_xml)
}

fn extract_from_xml(xml: &str) -> Result<TableTemplate> {
    let table_start = find_table_start(xml)
        .ok_or_else(|| Error::Template("No table found in table template".to_string()))?;

    let table_xml = extract_element(xml, table_start, "</w:tbl>")?;
    let rows = extract_rows(&table_xml);

    if rows.len() < 3 {
        return Err(Error::Template(
            "Table template must have at least 3 rows (header, odd, even)".to_string(),
        ));
    }

    let header = extract_row_style(&rows[0]);
    let row_odd = extract_row_style(&rows[1]);
    let row_even = extract_row_style(&rows[2]);

    // For first column vs other columns, we look at Row 3 if it exists, otherwise use Row 1
    let (first_column, other_columns) = if rows.len() >= 4 {
        let first = extract_cell_style(&rows[3], true);
        let other = extract_cell_style(&rows[3], false);
        (first, other)
    } else {
        let first = extract_cell_style(&rows[1], true);
        let other = extract_cell_style(&rows[1], false);
        (first, other)
    };

    let borders = extract_borders(&table_xml);
    let caption = extract_caption_style(xml, table_start);
    let cell_margins = extract_cell_margins(&table_xml);
    let cell_spacing = extract_cell_spacing(&table_xml);

    Ok(TableTemplate {
        header,
        row_odd,
        row_even,
        first_column,
        other_columns,
        borders,
        caption,
        cell_margins,
        cell_spacing,
    })
}

fn find_table_start(xml: &str) -> Option<usize> {
    xml.find("<w:tbl")
}

fn extract_element(xml: &str, start_pos: usize, close_tag: &str) -> Result<String> {
    let fragment = &xml[start_pos..];
    if let Some(end_pos) = fragment.find(close_tag) {
        Ok(fragment[..end_pos + close_tag.len()].to_string())
    } else {
        Err(Error::Template(format!(
            "Failed to find closing tag {}",
            close_tag
        )))
    }
}

fn extract_rows(table_xml: &str) -> Vec<String> {
    let mut rows = Vec::new();
    let mut pos = 0;
    while let Some(row_start) = table_xml[pos..].find("<w:tr") {
        let absolute_row_start = pos + row_start;
        if let Some(row_end) = table_xml[absolute_row_start..].find("</w:tr>") {
            rows.push(table_xml[absolute_row_start..absolute_row_start + row_end + 7].to_string());
            pos = absolute_row_start + row_end + 7;
        } else {
            break;
        }
    }
    rows
}

fn extract_row_style(row_xml: &str) -> RowStyle {
    let mut style = RowStyle::default();

    // Background color from first cell
    if let Some(cell_start) = row_xml.find("<w:tc") {
        if let Some(cell_end) = row_xml[cell_start..].find("</w:tc>") {
            let cell_xml = &row_xml[cell_start..cell_start + cell_end + 7];
            if let Some(shd) = extract_cell_shading(cell_xml) {
                style.background_color = Some(shd);
            }

            // Font properties from first run in the cell
            let (font, size, color, bold, italic) = extract_run_properties(cell_xml);
            style.font_family = font;
            style.font_size = size;
            style.font_color = color;
            style.bold = bold;
            style.italic = italic;
        }
    }

    style
}

fn extract_cells(row_xml: &str) -> Vec<String> {
    let mut cells = Vec::new();
    let mut pos = 0;
    while let Some(cell_start) = row_xml[pos..].find("<w:tc") {
        let absolute_cell_start = pos + cell_start;
        if let Some(cell_end) = row_xml[absolute_cell_start..].find("</w:tc>") {
            cells
                .push(row_xml[absolute_cell_start..absolute_cell_start + cell_end + 7].to_string());
            pos = absolute_cell_start + cell_end + 7;
        } else {
            break;
        }
    }
    cells
}

fn extract_cell_style(row_xml: &str, is_first_col: bool) -> CellStyle {
    let mut style = CellStyle::default();
    let cells = extract_cells(row_xml);

    let cell_xml = if is_first_col {
        cells.first()
    } else {
        cells.get(1).or_else(|| cells.first())
    };

    if let Some(cell_xml) = cell_xml {
        let (font, size, color, bold, italic) = extract_run_properties(cell_xml);
        style.font_family = font;
        style.font_size = size;
        style.font_color = color;
        style.bold = bold;
        style.italic = italic;

        // Alignment (look in paragraph properties inside cell)
        if let Some(p_start) = cell_xml.find("<w:p") {
            if let Some(p_end) = cell_xml[p_start..].find("</w:p>") {
                let p_xml = &cell_xml[p_start..p_start + p_end + 6];
                if let Some(jc) = extract_attribute(p_xml, "w:jc w:val=") {
                    style.alignment = jc;
                }
            }
        }

        // Vertical alignment (look in cell properties)
        if let Some(tc_pr_start) = cell_xml.find("<w:tcPr") {
            if let Some(tc_pr_end) = cell_xml[tc_pr_start..].find("</w:tcPr>") {
                let tc_pr_xml = &cell_xml[tc_pr_start..tc_pr_start + tc_pr_end + 9];
                if let Some(v_align) = extract_attribute(tc_pr_xml, "w:vAlign w:val=") {
                    style.vertical_alignment = v_align;
                }
            }
        }
    }

    style
}

fn extract_cell_shading(cell_xml: &str) -> Option<String> {
    if let Some(shd_pos) = cell_xml.find("<w:shd") {
        if let Some(fill) = extract_attribute(&cell_xml[shd_pos..], "w:fill=") {
            if fill == "auto" {
                return None;
            }
            return Some(format!("#{}", fill));
        }
    }
    None
}

fn extract_run_properties(xml: &str) -> (String, u32, String, bool, bool) {
    let mut font_family = "Calibri".to_string();
    let mut font_size = 22u32;
    let mut font_color = "#000000".to_string();
    let mut bold = false;
    let mut italic = false;

    if let Some(rpr_start) = xml.find("<w:rPr") {
        if let Some(rpr_end) = xml[rpr_start..].find("</w:rPr>") {
            let rpr_xml = &xml[rpr_start..rpr_start + rpr_end + 8];

            if rpr_xml.contains("<w:b/>") || rpr_xml.contains("<w:b ") {
                bold = true;
            }
            if rpr_xml.contains("<w:i/>") || rpr_xml.contains("<w:i ") {
                italic = true;
            }

            // Extract font size
            if let Some(pos) = rpr_xml.find("<w:szCs") {
                if let Some(val) = extract_attribute(&rpr_xml[pos..], "w:val=") {
                    if let Ok(s) = val.parse::<u32>() {
                        font_size = s;
                    }
                }
            } else if let Some(pos) = rpr_xml.find("<w:sz") {
                if let Some(val) = extract_attribute(&rpr_xml[pos..], "w:val=") {
                    if let Ok(s) = val.parse::<u32>() {
                        font_size = s;
                    }
                }
            }

            // Extract color
            if let Some(pos) = rpr_xml.find("<w:color") {
                if let Some(val) = extract_attribute(&rpr_xml[pos..], "w:val=") {
                    font_color = format!("#{}", val);
                }
            }

            // Extract font family
            if let Some(pos) = rpr_xml.find("<w:rFonts") {
                let fonts_xml = &rpr_xml[pos..];
                if let Some(font) = extract_attribute(fonts_xml, "w:cs=") {
                    font_family = font;
                } else if let Some(font) = extract_attribute(fonts_xml, "w:ascii=") {
                    font_family = font;
                } else if let Some(font) = extract_attribute(fonts_xml, "w:hAnsi=") {
                    font_family = font;
                }
            }
        }
    }

    (font_family, font_size, font_color, bold, italic)
}

fn extract_borders(table_xml: &str) -> BorderStyles {
    let mut borders = BorderStyles::default();

    if let Some(tbl_pr_pos) = table_xml.find("<w:tblPr") {
        if let Some(tbl_pr_end) = table_xml[tbl_pr_pos..].find("</w:tblPr>") {
            let tbl_pr = &table_xml[tbl_pr_pos..tbl_pr_pos + tbl_pr_end + 10];

            if let Some(borders_pos) = tbl_pr.find("<w:tblBorders") {
                if let Some(borders_end) = tbl_pr[borders_pos..].find("</w:tblBorders>") {
                    let borders_xml = &tbl_pr[borders_pos..borders_pos + borders_end + 15];

                    if let Some(top) = extract_border_style_tag(borders_xml, "<w:top") {
                        borders.top = top;
                    }
                    if let Some(bottom) = extract_border_style_tag(borders_xml, "<w:bottom") {
                        borders.bottom = bottom;
                    }
                    if let Some(left) = extract_border_style_tag(borders_xml, "<w:left") {
                        borders.left = left;
                    }
                    if let Some(right) = extract_border_style_tag(borders_xml, "<w:right") {
                        borders.right = right;
                    }
                    if let Some(inside_h) = extract_border_style_tag(borders_xml, "<w:insideH") {
                        borders.inside_h = inside_h;
                    }
                    if let Some(inside_v) = extract_border_style_tag(borders_xml, "<w:insideV") {
                        borders.inside_v = inside_v;
                    }
                }
            }
        }
    }

    borders
}

fn extract_border_style_tag(borders_xml: &str, tag: &str) -> Option<BorderStyle> {
    if let Some(pos) = borders_xml.find(tag) {
        if let Some(end) = borders_xml[pos..].find("/>") {
            let tag_xml = &borders_xml[pos..pos + end + 2];
            let mut style = BorderStyle::default();

            if let Some(val) = extract_attribute(tag_xml, "w:val=") {
                style.style = val;
            }
            if let Some(color) = extract_attribute(tag_xml, "w:color=") {
                style.color = format!("#{}", color);
            }
            if let Some(sz) = extract_attribute(tag_xml, "w:sz=") {
                if let Ok(s) = sz.parse::<u32>() {
                    style.width = s;
                }
            }
            return Some(style);
        }
    }
    None
}

fn extract_caption_style(xml: &str, table_start: usize) -> TableCaptionStyle {
    let mut style = TableCaptionStyle::default();

    // Search backwards from table_start for the nearest <w:p>
    let content_before = &xml[..table_start];
    if let Some(p_start) = content_before.rfind("<w:p") {
        let p_fragment = &xml[p_start..table_start];
        if let Some(p_end) = p_fragment.find("</w:p>") {
            let p_xml = &p_fragment[..p_end + 6];

            // Check if it's a caption placeholder
            if p_xml.contains("table_caption_prefix") {
                let (font, size, color, bold, italic) = extract_run_properties(p_xml);
                style.font_family = font;
                style.font_size = size;
                style.font_color = color;
                style.bold = bold;
                style.italic = italic;

                if let Some(jc) = extract_attribute(p_xml, "w:jc w:val=") {
                    style.alignment = jc;
                }

                // Extract spacing
                if let Some(spacing_pos) = p_xml.find("<w:spacing") {
                    if let Some(before) = extract_attribute(&p_xml[spacing_pos..], "w:before=") {
                        if let Ok(v) = before.parse::<u32>() {
                            style.spacing_before = v;
                        }
                    }
                    if let Some(after) = extract_attribute(&p_xml[spacing_pos..], "w:after=") {
                        if let Ok(v) = after.parse::<u32>() {
                            style.spacing_after = v;
                        }
                    }
                }

                // Try to extract prefix
                if let Some(t_start) = p_xml.find("<w:t") {
                    if let Some(t_end_tag) = p_xml[t_start..].find(">") {
                        let text_start = t_start + t_end_tag + 1;
                        if let Some(t_end) = p_xml[text_start..].find("</w:t>") {
                            let text = &p_xml[text_start..text_start + t_end];
                            if let Some(placeholder_start) = text.find("{{") {
                                let prefix = text[..placeholder_start].trim();
                                if !prefix.is_empty() {
                                    style.prefix = prefix.to_string();
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    style
}

fn extract_attribute(xml: &str, attr_name: &str) -> Option<String> {
    if let Some(pos) = xml.find(attr_name) {
        let start = pos + attr_name.len();
        let rest = &xml[start..];

        // Find the opening quote
        if let Some(quote_pos) = rest.find('"') {
            let after_quote = &rest[quote_pos + 1..];
            // Find the closing quote
            if let Some(end_quote) = after_quote.find('"') {
                return Some(after_quote[..end_quote].to_string());
            }
        }
    }
    None
}

fn extract_cell_margins(table_xml: &str) -> CellMargins {
    let mut margins = CellMargins::default();

    // Look for <w:tblCellMar> inside <w:tblPr>
    if let Some(mar_start) = table_xml.find("<w:tblCellMar") {
        if let Some(mar_end) = table_xml[mar_start..].find("</w:tblCellMar>") {
            let mar_xml = &table_xml[mar_start..mar_start + mar_end + 15];

            // Extract each margin: <w:top w:w="100" w:type="dxa"/>
            if let Some(val) = extract_margin_value(mar_xml, "w:top") {
                margins.top = val;
            }
            if let Some(val) = extract_margin_value(mar_xml, "w:bottom") {
                margins.bottom = val;
            }
            if let Some(val) = extract_margin_value(mar_xml, "w:left") {
                margins.left = val;
            }
            if let Some(val) = extract_margin_value(mar_xml, "w:right") {
                margins.right = val;
            }
        }
    }

    margins
}

fn extract_margin_value(xml: &str, tag: &str) -> Option<u32> {
    // Find <w:top w:w="100" .../>
    if let Some(pos) = xml.find(&format!("<{}", tag)) {
        let fragment = &xml[pos..];
        if let Some(w_val) = extract_attribute(fragment, "w:w=") {
            return w_val.parse().ok();
        }
    }
    None
}

fn extract_cell_spacing(table_xml: &str) -> CellSpacing {
    let mut spacing = CellSpacing::default();

    // Look for <w:spacing> in paragraph properties inside table
    // Check first cell's paragraph for spacing
    if let Some(spacing_start) = table_xml.find("<w:spacing") {
        let fragment = &table_xml[spacing_start..];

        if let Some(line) = extract_attribute(fragment, "w:line=") {
            spacing.line = line.parse().unwrap_or(240);
        }
        if let Some(rule) = extract_attribute(fragment, "w:lineRule=") {
            spacing.line_rule = rule;
        }
        if let Some(before) = extract_attribute(fragment, "w:before=") {
            spacing.before = before.parse().unwrap_or(0);
        }
        if let Some(after) = extract_attribute(fragment, "w:after=") {
            spacing.after = after.parse().unwrap_or(0);
        }
    }

    spacing
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
    fn test_cell_margins_default() {
        let margins = CellMargins::default();
        assert_eq!(margins.left, 108);
        assert_eq!(margins.right, 108);
        assert_eq!(margins.top, 0);
        assert_eq!(margins.bottom, 0);
    }

    #[test]
    fn test_cell_spacing_default() {
        let spacing = CellSpacing::default();
        assert_eq!(spacing.line, 240);
        assert_eq!(spacing.line_rule, "auto");
    }

    #[test]
    fn test_extract_file_not_found() {
        let result = extract(Path::new("/nonexistent/table.docx"));
        assert!(result.is_err());
    }

    #[test]
    fn test_debug_extract_real_template() {
        let path = Path::new(
            "/Users/kapong/workspace/lab/md2docx/examples/thai-manual/template/table.docx",
        );
        if !path.exists() {
            println!("Template file not found, skipping test");
            return;
        }

        let result = extract(path);
        match result {
            Ok(template) => {
                println!("\n=== TABLE TEMPLATE EXTRACTION DEBUG ===\n");

                println!("HEADER:");
                println!("  background_color: {:?}", template.header.background_color);
                println!("  font_family: {}", template.header.font_family);
                println!(
                    "  font_size: {} ({}pt)",
                    template.header.font_size,
                    template.header.font_size / 2
                );
                println!("  font_color: {}", template.header.font_color);
                println!("  bold: {}", template.header.bold);
                println!("  italic: {}", template.header.italic);

                println!("\nROW_ODD:");
                println!(
                    "  background_color: {:?}",
                    template.row_odd.background_color
                );
                println!("  font_family: {}", template.row_odd.font_family);
                println!(
                    "  font_size: {} ({}pt)",
                    template.row_odd.font_size,
                    template.row_odd.font_size / 2
                );
                println!("  font_color: {}", template.row_odd.font_color);
                println!("  bold: {}", template.row_odd.bold);
                println!("  italic: {}", template.row_odd.italic);

                println!("\nROW_EVEN:");
                println!(
                    "  background_color: {:?}",
                    template.row_even.background_color
                );
                println!("  font_family: {}", template.row_even.font_family);
                println!(
                    "  font_size: {} ({}pt)",
                    template.row_even.font_size,
                    template.row_even.font_size / 2
                );
                println!("  font_color: {}", template.row_even.font_color);
                println!("  bold: {}", template.row_even.bold);
                println!("  italic: {}", template.row_even.italic);

                println!("\nFIRST_COLUMN:");
                println!("  font_family: {}", template.first_column.font_family);
                println!(
                    "  font_size: {} ({}pt)",
                    template.first_column.font_size,
                    template.first_column.font_size / 2
                );
                println!("  font_color: {}", template.first_column.font_color);
                println!("  bold: {}", template.first_column.bold);
                println!("  italic: {}", template.first_column.italic);
                println!("  alignment: {}", template.first_column.alignment);
                println!(
                    "  vertical_alignment: {}",
                    template.first_column.vertical_alignment
                );

                println!("\nOTHER_COLUMNS:");
                println!("  font_family: {}", template.other_columns.font_family);
                println!(
                    "  font_size: {} ({}pt)",
                    template.other_columns.font_size,
                    template.other_columns.font_size / 2
                );
                println!("  font_color: {}", template.other_columns.font_color);
                println!("  bold: {}", template.other_columns.bold);
                println!("  italic: {}", template.other_columns.italic);
                println!("  alignment: {}", template.other_columns.alignment);
                println!(
                    "  vertical_alignment: {}",
                    template.other_columns.vertical_alignment
                );

                println!("\nBORDERS:");
                println!(
                    "  top: style={}, color={}, width={}",
                    template.borders.top.style,
                    template.borders.top.color,
                    template.borders.top.width
                );
                println!(
                    "  bottom: style={}, color={}, width={}",
                    template.borders.bottom.style,
                    template.borders.bottom.color,
                    template.borders.bottom.width
                );
                println!(
                    "  inside_h: style={}, color={}, width={}",
                    template.borders.inside_h.style,
                    template.borders.inside_h.color,
                    template.borders.inside_h.width
                );
                println!(
                    "  inside_v: style={}, color={}, width={}",
                    template.borders.inside_v.style,
                    template.borders.inside_v.color,
                    template.borders.inside_v.width
                );

                println!("\nCELL_MARGINS:");
                println!(
                    "  top={}, bottom={}, left={}, right={}",
                    template.cell_margins.top,
                    template.cell_margins.bottom,
                    template.cell_margins.left,
                    template.cell_margins.right
                );

                println!("\nCELL_SPACING:");
                println!(
                    "  line={}, line_rule={}, before={}, after={}",
                    template.cell_spacing.line,
                    template.cell_spacing.line_rule,
                    template.cell_spacing.before,
                    template.cell_spacing.after
                );

                println!("\nCAPTION:");
                println!("  prefix: {}", template.caption.prefix);
                println!("  font_family: {}", template.caption.font_family);
                println!("  bold: {}", template.caption.bold);
                println!("  italic: {}", template.caption.italic);

                println!("\n=== END DEBUG ===\n");
            }
            Err(e) => {
                println!("Error extracting template: {:?}", e);
            }
        }
    }
}
