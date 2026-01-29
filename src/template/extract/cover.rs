//! Cover template extraction from DOCX files
//!
//! Extracts cover page design elements from a DOCX file created in Microsoft Word.
//! The cover page can contain shapes, images, text boxes with placeholders.

use crate::error::{Error, Result};
use std::io::Read;
use std::path::Path;

/// Represents an extracted cover page template
#[derive(Debug, Clone)]
pub struct CoverTemplate {
    /// Background color of the page (hex, e.g., "#FFFFFF")
    pub background_color: Option<String>,
    /// Elements on the cover page (shapes, images, text)
    pub elements: Vec<CoverElement>,
    /// Page width in twips
    pub page_width: u32,
    /// Page height in twips
    pub page_height: u32,
    /// Page margins in twips
    pub margins: PageMargins,
    /// Raw XML content of the cover page (for direct copying)
    pub raw_xml: Option<String>,
}

/// Page margins
#[derive(Debug, Clone, Copy)]
pub struct PageMargins {
    pub top: u32,
    pub right: u32,
    pub bottom: u32,
    pub left: u32,
}

impl Default for PageMargins {
    fn default() -> Self {
        Self {
            top: 1440, // 1 inch
            right: 1440,
            bottom: 1440,
            left: 1440,
        }
    }
}

/// Individual element on cover page
#[derive(Debug, Clone)]
pub enum CoverElement {
    /// Text element with placeholder support
    Text {
        /// Text content (may contain {{placeholder}})
        content: String,
        /// X position in EMUs (English Metric Units)
        x: i64,
        /// Y position in EMUs
        y: i64,
        /// Width in EMUs (for text wrapping)
        width: i64,
        /// Height in EMUs
        height: i64,
        /// Font family name
        font_family: String,
        /// Font size in half-points
        font_size: u32,
        /// Text color (hex, e.g., "#1a365d")
        color: String,
        /// Whether text is bold
        bold: bool,
        /// Whether text is italic
        italic: bool,
        /// Text alignment: "left", "center", "right", "both"
        alignment: String,
    },
    /// Shape element (rectangle, line, etc.)
    Shape {
        /// Type of shape
        shape_type: ShapeType,
        /// X position in EMUs
        x: i64,
        /// Y position in EMUs
        y: i64,
        /// Width in EMUs
        width: i64,
        /// Height in EMUs
        height: i64,
        /// Fill color (hex, e.g., "#1a365d")
        fill_color: Option<String>,
        /// Stroke/border color (hex)
        stroke_color: Option<String>,
        /// Stroke width in EMUs
        stroke_width: u32,
    },
    /// Image element
    Image {
        /// Relationship ID for the image
        rel_id: String,
        /// X position in EMUs
        x: i64,
        /// Y position in EMUs
        y: i64,
        /// Width in EMUs
        width: i64,
        /// Height in EMUs
        height: i64,
        /// Image filename (e.g., "image1.png")
        filename: String,
        /// Image data bytes (loaded from cover.docx)
        data: Option<Vec<u8>>,
    },
}

/// Type of shape
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Rectangle shape
    Rectangle,
    /// Line shape
    Line,
    /// Circle/ellipse shape
    Circle,
}

impl Default for CoverTemplate {
    fn default() -> Self {
        Self {
            background_color: None,
            elements: Vec::new(),
            page_width: 11906,  // A4 width in twips
            page_height: 16838, // A4 height in twips
            margins: PageMargins::default(),
            raw_xml: None,
        }
    }
}

impl CoverTemplate {
    /// Create a new empty cover template
    pub fn new() -> Self {
        Self::default()
    }

    /// Add an element to the cover
    pub fn add_element(mut self, element: CoverElement) -> Self {
        self.elements.push(element);
        self
    }

    /// Set background color
    pub fn with_background_color(mut self, color: impl Into<String>) -> Self {
        self.background_color = Some(color.into());
        self
    }

    /// Set page size
    pub fn with_page_size(mut self, width: u32, height: u32) -> Self {
        self.page_width = width;
        self.page_height = height;
        self
    }

    /// Set page margins
    pub fn with_margins(mut self, top: u32, right: u32, bottom: u32, left: u32) -> Self {
        self.margins = PageMargins {
            top,
            right,
            bottom,
            left,
        };
        self
    }

    /// Check if the cover has any elements
    pub fn is_empty(&self) -> bool {
        self.elements.is_empty() && self.raw_xml.is_none()
    }

    /// Get all text elements that contain placeholders
    pub fn text_elements_with_placeholders(&self) -> Vec<&CoverElement> {
        self.elements
            .iter()
            .filter(|e| {
                if let CoverElement::Text { content, .. } = e {
                    content.contains("{{") && content.contains("}}")
                } else {
                    false
                }
            })
            .collect()
    }

    /// Check if this template has raw XML that should be used directly
    pub fn has_raw_xml(&self) -> bool {
        self.raw_xml.is_some()
    }
}

/// Extract cover template from a DOCX file
///
/// This function reads a DOCX file and extracts:
/// - Page background color
/// - Text elements with their formatting and positions
/// - Shape elements (rectangles, lines, circles)
/// - Image elements
///
/// # Arguments
/// * `path` - Path to the cover.docx file
///
/// # Returns
/// The extracted `CoverTemplate`
pub fn extract(path: &Path) -> Result<CoverTemplate> {
    if !path.exists() {
        return Err(Error::Template(format!(
            "Cover template file not found: {}",
            path.display()
        )));
    }

    // Read the DOCX file (it's a ZIP archive)
    let file = std::fs::File::open(path)?;
    let mut archive = zip::ZipArchive::new(file)?;

    // Read document.xml
    let mut document_xml = String::new();
    {
        let mut doc_file = archive
            .by_name("word/document.xml")
            .map_err(|e| Error::Template(format!("Failed to read document.xml: {}", e)))?;
        doc_file.read_to_string(&mut document_xml)?;
    }

    // Read relationships to map rId to filenames
    let rels_xml =
        read_archive_file(&mut archive, "word/_rels/document.xml.rels").unwrap_or_default();

    // Extract the first page content (before first sectPr or page break)
    let cover_xml = extract_first_page(&document_xml);

    // Parse page properties
    let (width, height, margins) = parse_page_properties(&document_xml);

    // Parse elements from the cover XML
    let mut elements = parse_cover_elements(&cover_xml)?;

    // Load image data for image elements
    for element in &mut elements {
        if let CoverElement::Image {
            rel_id,
            filename,
            data,
            ..
        } = element
        {
            // Find the actual filename from relationships
            if let Some(img_path) = find_image_path_from_rel_id(&rels_xml, rel_id) {
                // Use just the filename part to avoid directory issues
                let file_name = std::path::Path::new(&img_path)
                    .file_name()
                    .and_then(|n| n.to_str())
                    .unwrap_or("image.png")
                    .to_string();
                *filename = file_name;

                // Try to load the image data using the full path from the archive
                // Relationships usually point to "media/image1.png" which is relative to "word/"
                // So full path is "word/media/image1.png"
                let full_path = format!("word/{}", img_path);
                if let Ok(img_data) = read_archive_file_bytes(&mut archive, &full_path) {
                    *data = Some(img_data);
                }
            }
        }
    }

    // Extract background color if any
    let background_color = extract_background_color(&document_xml);

    Ok(CoverTemplate {
        background_color,
        elements,
        page_width: width,
        page_height: height,
        margins,
        raw_xml: Some(cover_xml),
    })
}

/// Read a file from the ZIP archive as string
fn read_archive_file(archive: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Result<String> {
    let mut content = String::new();
    if let Ok(mut file) = archive.by_name(name) {
        file.read_to_string(&mut content)?;
    }
    Ok(content)
}

/// Read a file from the ZIP archive as bytes
fn read_archive_file_bytes(
    archive: &mut zip::ZipArchive<std::fs::File>,
    name: &str,
) -> Result<Vec<u8>> {
    let mut content = Vec::new();
    if let Ok(mut file) = archive.by_name(name) {
        file.read_to_end(&mut content)?;
    }
    Ok(content)
}

/// Find image path from relationship ID
fn find_image_path_from_rel_id(rels_xml: &str, rel_id: &str) -> Option<String> {
    // Look for Relationship with Id="rel_id" and get Target
    let search_pattern = format!(r#"Id="{}""#, rel_id);
    if let Some(pos) = rels_xml.find(&search_pattern) {
        // Look at next 200 chars, or rest of string if shorter
        let end_pos = (pos + 200).min(rels_xml.len());
        let rel_fragment = &rels_xml[pos..end_pos];
        if let Some(target) = extract_attribute(rel_fragment, "Target=") {
            return Some(target);
        }
    }
    None
}

/// Extract the first page content from document XML
fn extract_first_page(document_xml: &str) -> String {
    // Find the body content
    if let Some(body_start) = document_xml.find("<w:body>") {
        if let Some(body_end) = document_xml.find("</w:body>") {
            let body_content = &document_xml[body_start + 8..body_end];

            // Find the first section break or return all content
            if let Some(sect_pr_pos) = body_content.find("<w:sectPr") {
                // Check if this sectPr is inside a paragraph
                // Look for the start of the enclosing paragraph
                let content_before = &body_content[..sect_pr_pos];
                // Check for <w:p> or <w:p ...
                let p_start_1 = content_before.rfind("<w:p>");
                let p_start_2 = content_before.rfind("<w:p ");
                let p_start = match (p_start_1, p_start_2) {
                    (Some(a), Some(b)) => Some(std::cmp::max(a, b)),
                    (Some(a), None) => Some(a),
                    (None, Some(b)) => Some(b),
                    (None, None) => None,
                };

                let p_end = content_before.rfind("</w:p>");

                let cut_pos = match (p_start, p_end) {
                    // Start of para is after the last end of para -> we are inside a para
                    (Some(start), Some(end)) if start > end => start,
                    // Start of para found, no end of para -> we are inside the first para
                    (Some(start), None) => start,
                    // Otherwise, we are likely at body level (or parsing failed logic)
                    _ => sect_pr_pos,
                };

                body_content[..cut_pos].to_string()
            } else {
                body_content.to_string()
            }
        } else {
            document_xml.to_string()
        }
    } else {
        document_xml.to_string()
    }
}

/// Parse page properties from document XML
fn parse_page_properties(document_xml: &str) -> (u32, u32, PageMargins) {
    let mut width = 11906u32; // A4 default
    let mut height = 16838u32;
    let mut margins = PageMargins::default();

    // Parse page size
    if let Some(pg_sz_start) = document_xml.find("<w:pgSz") {
        if let Some(pg_sz_end) = document_xml[pg_sz_start..].find(">") {
            let pg_sz = &document_xml[pg_sz_start..pg_sz_start + pg_sz_end + 1];

            if let Some(w_attr) = extract_attribute(pg_sz, "w:w=") {
                if let Ok(w) = w_attr.parse::<u32>() {
                    width = w;
                }
            }
            if let Some(h_attr) = extract_attribute(pg_sz, "w:h=") {
                if let Ok(h) = h_attr.parse::<u32>() {
                    height = h;
                }
            }
        }
    }

    // Parse margins
    if let Some(pg_mar_start) = document_xml.find("<w:pgMar") {
        if let Some(pg_mar_end) = document_xml[pg_mar_start..].find(">") {
            let pg_mar = &document_xml[pg_mar_start..pg_mar_start + pg_mar_end + 1];

            if let Some(top) = extract_attribute(pg_mar, "w:top=") {
                if let Ok(t) = top.parse::<u32>() {
                    margins.top = t;
                }
            }
            if let Some(right) = extract_attribute(pg_mar, "w:right=") {
                if let Ok(r) = right.parse::<u32>() {
                    margins.right = r;
                }
            }
            if let Some(bottom) = extract_attribute(pg_mar, "w:bottom=") {
                if let Ok(b) = bottom.parse::<u32>() {
                    margins.bottom = b;
                }
            }
            if let Some(left) = extract_attribute(pg_mar, "w:left=") {
                if let Ok(l) = left.parse::<u32>() {
                    margins.left = l;
                }
            }
        }
    }

    (width, height, margins)
}

/// Extract attribute value from XML element
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

/// Extract background color from document XML
fn extract_background_color(document_xml: &str) -> Option<String> {
    // Look for document background
    if let Some(bg_start) = document_xml.find("<w:background") {
        if let Some(color_attr) = extract_attribute(&document_xml[bg_start..], "w:color=") {
            return Some(format!("#{}", color_attr));
        }
    }
    None
}

/// Parse cover elements from XML
fn parse_cover_elements(cover_xml: &str) -> Result<Vec<CoverElement>> {
    let mut elements = Vec::new();

    // Parse paragraphs (text elements)
    let mut pos = 0;
    while let Some(p_start) = cover_xml[pos..].find("<w:p") {
        let absolute_p_start = pos + p_start;
        if let Some(p_end) = cover_xml[absolute_p_start..].find("</w:p>") {
            let p_xml = &cover_xml[absolute_p_start..absolute_p_start + p_end + 6];

            if let Some(element) = parse_paragraph_element(p_xml)? {
                elements.push(element);
            }

            pos = absolute_p_start + p_end + 6;
        } else {
            break;
        }
    }

    // Parse drawings (shapes and images)
    pos = 0;
    while let Some(drawing_start) = cover_xml[pos..].find("<w:drawing>") {
        let absolute_drawing_start = pos + drawing_start;
        if let Some(drawing_end) = cover_xml[absolute_drawing_start..].find("</w:drawing>") {
            let drawing_xml =
                &cover_xml[absolute_drawing_start..absolute_drawing_start + drawing_end + 12];

            if let Some(element) = parse_drawing_element(drawing_xml)? {
                elements.push(element);
            }

            pos = absolute_drawing_start + drawing_end + 12;
        } else {
            break;
        }
    }

    Ok(elements)
}

/// Parse a paragraph element
fn parse_paragraph_element(p_xml: &str) -> Result<Option<CoverElement>> {
    // Extract text content
    let mut content = String::new();
    let mut pos = 0;

    while let Some(t_start) = p_xml[pos..].find("<w:t") {
        let absolute_t_start = pos + t_start;

        // Find the end of the opening tag
        if let Some(tag_end) = p_xml[absolute_t_start..].find(">") {
            let content_start = absolute_t_start + tag_end + 1;

            // Find the closing tag
            if let Some(t_end) = p_xml[content_start..].find("</w:t>") {
                let text = &p_xml[content_start..content_start + t_end];
                content.push_str(text);
                pos = content_start + t_end + 6;
            } else {
                break;
            }
        } else {
            break;
        }
    }

    if content.is_empty() {
        return Ok(None);
    }

    // Extract run properties for formatting
    let (font_family, font_size, color, bold, italic) = parse_run_properties(p_xml);

    // Extract paragraph properties for alignment
    let alignment = parse_paragraph_alignment(p_xml);

    // For now, position is estimated - in a full implementation,
    // we'd parse the actual positioning from text boxes or frames
    Ok(Some(CoverElement::Text {
        content,
        x: 0,
        y: 0,
        width: 6000000, // Default width in EMUs (~6 inches)
        height: 200000, // Default height in EMUs
        font_family,
        font_size,
        color,
        bold,
        italic,
        alignment,
    }))
}

/// Parse run properties (formatting)
fn parse_run_properties(xml: &str) -> (String, u32, String, bool, bool) {
    let mut font_family = "TH Sarabun New".to_string();
    let mut font_size = 24u32; // 12pt default
    let mut color = "#000000".to_string();
    let mut bold = false;
    let mut italic = false;

    // Check for bold
    if xml.contains("<w:b/>") || xml.contains("<w:b ") {
        bold = true;
    }

    // Check for italic
    if xml.contains("<w:i/>") || xml.contains("<w:i ") {
        italic = true;
    }

    // Extract font size
    if let Some(sz) = extract_attribute(xml, "<w:sz w:val=") {
        if let Ok(size) = sz.parse::<u32>() {
            font_size = size;
        }
    }

    // Extract color
    if let Some(color_attr) = extract_attribute(xml, "<w:color w:val=") {
        color = format!("#{}", color_attr);
    }

    // Extract font family
    if let Some(fonts_attr) = extract_attribute(xml, "w:ascii=") {
        font_family = fonts_attr;
    }

    (font_family, font_size, color, bold, italic)
}

/// Parse paragraph alignment
fn parse_paragraph_alignment(xml: &str) -> String {
    if let Some(jc) = extract_attribute(xml, "<w:jc w:val=") {
        jc
    } else {
        "left".to_string()
    }
}

/// Parse a drawing element (shapes or images)
fn parse_drawing_element(drawing_xml: &str) -> Result<Option<CoverElement>> {
    // Check if it's an image
    if drawing_xml.contains("<a:blip") || drawing_xml.contains("pic:pic") {
        // Parse image
        let x = extract_emu_value(drawing_xml, "x=").unwrap_or(0);
        let y = extract_emu_value(drawing_xml, "y=").unwrap_or(0);
        let width = extract_emu_value(drawing_xml, "cx=").unwrap_or(1000000);
        let height = extract_emu_value(drawing_xml, "cy=").unwrap_or(1000000);

        // Extract relationship ID
        let rel_id = if let Some(r_id) = extract_attribute(drawing_xml, "r:embed=") {
            r_id
        } else {
            return Ok(None);
        };

        return Ok(Some(CoverElement::Image {
            rel_id,
            x,
            y,
            width,
            height,
            filename: String::new(), // Will be filled in later from relationships
            data: None,              // Will be loaded later from archive
        }));
    }

    // Check if it's a shape
    if drawing_xml.contains("<a:rect") || drawing_xml.contains("<a:ellipse") {
        // Parse shape - simplified implementation
        let shape_type = if drawing_xml.contains("<a:ellipse") {
            ShapeType::Circle
        } else {
            ShapeType::Rectangle
        };

        let x = extract_emu_value(drawing_xml, "x=").unwrap_or(0);
        let y = extract_emu_value(drawing_xml, "y=").unwrap_or(0);
        let width = extract_emu_value(drawing_xml, "cx=").unwrap_or(1000000);
        let height = extract_emu_value(drawing_xml, "cy=").unwrap_or(1000000);

        // Extract fill color
        let fill_color = if drawing_xml.contains("<a:solidFill>") {
            extract_attribute(drawing_xml, "val=").map(|srgb| format!("#{}", srgb))
        } else {
            None
        };

        return Ok(Some(CoverElement::Shape {
            shape_type,
            x,
            y,
            width,
            height,
            fill_color,
            stroke_color: None,
            stroke_width: 0,
        }));
    }

    Ok(None)
}

/// Extract EMU value from XML
fn extract_emu_value(xml: &str, attr: &str) -> Option<i64> {
    extract_attribute(xml, attr).and_then(|v| v.parse::<i64>().ok())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cover_template_default() {
        let cover = CoverTemplate::default();
        assert!(cover.background_color.is_none());
        assert!(cover.elements.is_empty());
        assert_eq!(cover.page_width, 11906);
        assert_eq!(cover.page_height, 16838);
    }

    #[test]
    fn test_cover_template_builder() {
        let cover = CoverTemplate::new()
            .with_background_color("#FFFFFF")
            .with_page_size(11906, 16838)
            .with_margins(1440, 1440, 1440, 1440)
            .add_element(CoverElement::Text {
                content: "{{title}}".to_string(),
                x: 0,
                y: 0,
                width: 1000000,
                height: 200000,
                font_family: "Calibri".to_string(),
                font_size: 48,
                color: "#000000".to_string(),
                bold: true,
                italic: false,
                alignment: "center".to_string(),
            });

        assert_eq!(cover.background_color, Some("#FFFFFF".to_string()));
        assert_eq!(cover.elements.len(), 1);
        assert!(!cover.is_empty());
    }

    #[test]
    fn test_text_elements_with_placeholders() {
        let cover = CoverTemplate::new()
            .add_element(CoverElement::Text {
                content: "{{title}}".to_string(),
                x: 0,
                y: 0,
                width: 1000000,
                height: 200000,
                font_family: "Calibri".to_string(),
                font_size: 48,
                color: "#000000".to_string(),
                bold: true,
                italic: false,
                alignment: "center".to_string(),
            })
            .add_element(CoverElement::Text {
                content: "Static text".to_string(),
                x: 0,
                y: 200000,
                width: 1000000,
                height: 200000,
                font_family: "Calibri".to_string(),
                font_size: 24,
                color: "#000000".to_string(),
                bold: false,
                italic: false,
                alignment: "center".to_string(),
            });

        let with_placeholders = cover.text_elements_with_placeholders();
        assert_eq!(with_placeholders.len(), 1);
    }

    #[test]
    fn test_extract_file_not_found() {
        let result = extract(Path::new("/nonexistent/cover.docx"));
        assert!(result.is_err());
    }

    #[test]
    fn test_extract_attribute() {
        let xml = r#"<w:pgSz w:w="12240" w:h="15840"/>"#;
        assert_eq!(extract_attribute(xml, "w:w="), Some("12240".to_string()));
        assert_eq!(extract_attribute(xml, "w:h="), Some("15840".to_string()));
        assert_eq!(extract_attribute(xml, "w:notexist="), None);
    }

    #[test]
    fn test_parse_page_properties() {
        let xml = r#"
            <w:sectPr>
                <w:pgSz w:w="12240" w:h="15840"/>
                <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>
        "#;

        let (width, height, margins) = parse_page_properties(xml);
        assert_eq!(width, 12240);
        assert_eq!(height, 15840);
        assert_eq!(margins.top, 1440);
        assert_eq!(margins.left, 1440);
    }

    #[test]
    fn test_extract_first_page() {
        let xml = r#"<w:body>
            <w:p>First paragraph</w:p>
            <w:p>Second paragraph</w:p>
            <w:sectPr>Section properties</w:sectPr>
            <w:p>Third paragraph</w:p>
        </w:body>"#;

        let first_page = extract_first_page(xml);
        assert!(first_page.contains("First paragraph"));
        assert!(first_page.contains("Second paragraph"));
        assert!(!first_page.contains("Third paragraph"));
    }
}
