//! Template generation for dump-template command

use std::collections::HashSet;
use std::path::Path;

use crate::docx::ooxml::Language;
use crate::error::Result;
use crate::Document;

/// Required styles for md2docx to work correctly
pub const REQUIRED_STYLES: &[&str] = &[
    "Title",
    "Heading1",
    "Heading2",
    "Heading3",
    "Normal",
    "Code",
    "Quote",
    "ListParagraph",
];

/// Optional but recommended styles
pub const RECOMMENDED_STYLES: &[&str] = &[
    "Subtitle",
    "Heading4",
    "CodeChar",
    "Caption",
    "TOC1",
    "TOC2",
    "TOC3",
    "FootnoteText",
    "Hyperlink",
    "CodeFilename",
];

/// Result of template validation
#[derive(Debug, Default)]
pub struct ValidationResult {
    /// Styles that are missing and required
    pub missing_required: Vec<String>,
    /// Styles that are missing but only recommended
    pub missing_recommended: Vec<String>,
    /// Styles found in the template
    pub found_styles: Vec<String>,
    /// Warnings or notes
    pub warnings: Vec<String>,
}

impl ValidationResult {
    /// Returns true if all required styles are present
    pub fn is_valid(&self) -> bool {
        self.missing_required.is_empty()
    }
}

/// Validate a template DOCX file for required styles
///
/// Opens the DOCX file (which is a ZIP), extracts and parses word/styles.xml,
/// and checks if all required styles are present.
///
/// # Arguments
/// * `path` - Path to the template DOCX file
///
/// # Returns
/// A `Result` containing the ValidationResult with missing styles and warnings
///
/// # Example
/// ```rust,no_run
/// use md2docx::docx::template::validate_template;
/// use std::path::Path;
///
/// let result = validate_template(Path::new("template.docx")).unwrap();
/// if result.is_valid() {
///     println!("Template is valid!");
/// } else {
///     println!("Missing styles: {:?}", result.missing_required);
/// }
/// ```
#[cfg(not(target_arch = "wasm32"))]
pub fn validate_template(path: &Path) -> Result<ValidationResult> {
    use std::io::Read;
    use zip::ZipArchive;

    let file = std::fs::File::open(path)?;
    let mut archive = ZipArchive::new(file)?;

    let mut result = ValidationResult::default();
    let mut found_styles = HashSet::new();

    // Read word/styles.xml from the archive
    if let Ok(mut styles_file) = archive.by_name("word/styles.xml") {
        let mut content = String::new();
        styles_file.read_to_string(&mut content)?;

        // Use quick_xml for proper parsing
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(&content);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                    if e.name().as_ref() == b"w:style" =>
                {
                    // Extract styleId attribute
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"w:styleId" {
                            if let Ok(value) = std::str::from_utf8(&attr.value) {
                                found_styles.insert(value.to_string());
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    result.warnings.push(format!("XML parsing error: {}", e));
                    break;
                }
                _ => {}
            }
            buf.clear();
        }
    } else {
        result
            .warnings
            .push("Could not find word/styles.xml in template".to_string());
    }

    // Check required styles
    for style in REQUIRED_STYLES {
        if found_styles.contains(*style) {
            result.found_styles.push(style.to_string());
        } else {
            result.missing_required.push(style.to_string());
        }
    }

    // Check recommended styles
    for style in RECOMMENDED_STYLES {
        if !found_styles.contains(*style) {
            result.missing_recommended.push(style.to_string());
        }
    }

    Ok(result)
}

/// Generate a template DOCX with all required styles
///
/// The template includes sample content for each style so users can:
/// 1. Open in Word
/// 2. Modify styles (right-click style -> Modify)
/// 3. Save and use as template with md2docx
pub fn generate_template(lang: Language, minimal: bool) -> Result<Vec<u8>> {
    let mut doc = Document::with_language(lang);

    // Title
    doc = doc.add_styled_paragraph("Title", sample_text("Title", lang));

    // Subtitle
    doc = doc.add_styled_paragraph("Subtitle", sample_text("Subtitle", lang));

    // Headings
    doc = doc.add_heading(1, sample_text("Heading 1", lang));
    doc = doc.add_paragraph(sample_text("Normal paragraph", lang));

    doc = doc.add_heading(2, sample_text("Heading 2", lang));
    doc = doc.add_paragraph(sample_text("Normal paragraph", lang));

    doc = doc.add_heading(3, sample_text("Heading 3", lang));
    doc = doc.add_paragraph(sample_text("Normal paragraph", lang));

    if !minimal {
        doc = doc.add_heading(4, sample_text("Heading 4", lang));
        doc = doc.add_paragraph(sample_text("Normal paragraph", lang));
    }

    // Quote
    doc = doc.add_quote(sample_text("Blockquote", lang));

    // Code
    doc = doc.add_code_block("fn main() {\n    println!(\"Hello, World!\");\n}");

    // Caption (for figures/tables)
    doc = doc.add_styled_paragraph("Caption", sample_text("Figure caption", lang));

    if !minimal {
        // List (simulate with ListParagraph style)
        doc = doc.add_styled_paragraph("ListParagraph", "• List item 1");
        doc = doc.add_styled_paragraph("ListParagraph", "• List item 2");

        // TOC styles
        doc = doc.add_styled_paragraph("TOC1", "Table of Contents Level 1");
        doc = doc.add_styled_paragraph("TOC2", "    Table of Contents Level 2");
        doc = doc.add_styled_paragraph("TOC3", "        Table of Contents Level 3");

        // Footnote
        doc = doc.add_styled_paragraph("FootnoteText", sample_text("Footnote text", lang));
    }

    doc.to_bytes()
}

/// Get sample text for a style, localized if Thai
fn sample_text(style: &str, lang: Language) -> &'static str {
    match (style, lang) {
        ("Title", Language::Thai) => "ชื่อเอกสาร (Document Title)",
        ("Title", _) => "Document Title",

        ("Subtitle", Language::Thai) => "ชื่อรอง (Subtitle)",
        ("Subtitle", _) => "Document Subtitle",

        ("Heading 1", Language::Thai) => "หัวข้อ 1 (Heading 1)",
        ("Heading 1", _) => "Heading 1 Example",

        ("Heading 2", Language::Thai) => "หัวข้อ 2 (Heading 2)",
        ("Heading 2", _) => "Heading 2 Example",

        ("Heading 3", Language::Thai) => "หัวข้อ 3 (Heading 3)",
        ("Heading 3", _) => "Heading 3 Example",

        ("Heading 4", Language::Thai) => "หัวข้อ 4 (Heading 4)",
        ("Heading 4", _) => "Heading 4 Example",

        ("Normal paragraph", Language::Thai) => {
            "นี่คือย่อหน้าปกติ ใช้สำหรับเนื้อหาหลักของเอกสาร This is normal paragraph text for document content."
        }
        ("Normal paragraph", _) => {
            "This is a normal paragraph. You can modify this style in Word to change the default body text formatting."
        }

        ("Blockquote", Language::Thai) => "นี่คือข้อความอ้างอิง (This is a blockquote)",
        ("Blockquote", _) => {
            "This is a blockquote. Modify this style to change how quoted text appears."
        }

        ("Figure caption", Language::Thai) => "รูปที่ 1: คำอธิบายรูป (Figure 1: Caption)",
        ("Figure caption", _) => "Figure 1: This is a figure caption",

        ("Footnote text", Language::Thai) => "นี่คือข้อความเชิงอรรถ (Footnote text)",
        ("Footnote text", _) => "This is footnote text.",

        _ => "Sample text",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_required_styles_defined() {
        assert!(!REQUIRED_STYLES.is_empty());
        assert!(REQUIRED_STYLES.contains(&"Heading1"));
        assert!(REQUIRED_STYLES.contains(&"Normal"));
    }

    #[test]
    fn test_recommended_styles_defined() {
        assert!(!RECOMMENDED_STYLES.is_empty());
        assert!(RECOMMENDED_STYLES.contains(&"TOC1"));
        assert!(RECOMMENDED_STYLES.contains(&"CodeFilename"));
    }

    #[test]
    fn test_validation_result_default() {
        let result = ValidationResult::default();
        assert!(result.is_valid());
        assert!(result.missing_required.is_empty());
        assert!(result.missing_recommended.is_empty());
        assert!(result.found_styles.is_empty());
        assert!(result.warnings.is_empty());
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_validate_generated_template() {
        use std::io::Write;

        let temp_dir = std::env::temp_dir();
        let template_path = temp_dir.join("test_template.docx");

        // Generate a template
        let bytes = generate_template(Language::English, false).unwrap();

        // Write to temp file
        {
            let mut file = std::fs::File::create(&template_path).unwrap();
            file.write_all(&bytes).unwrap();
        }

        // Validate it
        let result = validate_template(&template_path).unwrap();

        // Should be valid (all required styles present)
        assert!(result.is_valid());
        assert!(result.missing_required.is_empty());

        // Clean up
        std::fs::remove_file(&template_path).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_validate_template_missing_styles() {
        use std::io::Write;
        use zip::write::{FileOptions, ZipWriter};

        let temp_dir = std::env::temp_dir();
        let template_path = temp_dir.join("invalid_template.docx");

        // Create a minimal DOCX with incomplete styles
        {
            let file = std::fs::File::create(&template_path).unwrap();
            let mut zip = ZipWriter::new(file);

            // Add minimal required structure
            let options: FileOptions<'static, ()> =
                FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

            // Add minimal styles.xml with only "Normal" style
            let styles_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>"#;

            zip.start_file("word/styles.xml", options).unwrap();
            zip.write_all(styles_content.as_bytes()).unwrap();

            // Add minimal document.xml
            let doc_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:t>Test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

            zip.start_file("word/document.xml", options).unwrap();
            zip.write_all(doc_content.as_bytes()).unwrap();

            // Add minimal content_types.xml
            let types_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

            zip.start_file("[Content_Types].xml", options).unwrap();
            zip.write_all(types_content.as_bytes()).unwrap();

            // Add minimal _rels/.rels
            let rels_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

            zip.start_file("_rels/.rels", options).unwrap();
            zip.write_all(rels_content.as_bytes()).unwrap();

            // Add word/_rels/document.xml.rels
            let doc_rels_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

            zip.start_file("word/_rels/document.xml.rels", options)
                .unwrap();
            zip.write_all(doc_rels_content.as_bytes()).unwrap();

            zip.finish().unwrap();
        }

        // Validate it
        let result = validate_template(&template_path).unwrap();

        // Should be invalid (missing Heading1, Heading2, Heading3, etc.)
        assert!(!result.is_valid());
        assert!(!result.missing_required.is_empty());
        assert!(result.missing_required.contains(&"Heading1".to_string()));

        // Clean up
        std::fs::remove_file(&template_path).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_validate_nonexistent_file() {
        use std::path::Path;

        let result = validate_template(Path::new("/nonexistent/template.docx"));
        assert!(result.is_err());
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_validate_invalid_zip() {
        use std::io::Write;

        let temp_dir = std::env::temp_dir();
        let invalid_path = temp_dir.join("invalid.docx");

        // Create a file that's not a valid ZIP
        {
            let mut file = std::fs::File::create(&invalid_path).unwrap();
            file.write_all(b"Not a valid DOCX file").unwrap();
        }

        let result = validate_template(&invalid_path);
        assert!(result.is_err());

        // Clean up
        std::fs::remove_file(&invalid_path).unwrap();
    }

    #[test]
    fn test_generate_template_english() {
        let bytes = generate_template(Language::English, false).unwrap();

        // Should be a valid ZIP
        assert!(!bytes.is_empty());
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_generate_template_thai() {
        let bytes = generate_template(Language::Thai, false).unwrap();

        // Should be a valid ZIP
        assert!(!bytes.is_empty());
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_generate_template_minimal() {
        let bytes = generate_template(Language::English, true).unwrap();

        // Should be a valid ZIP
        assert!(!bytes.is_empty());
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_sample_text_english() {
        assert_eq!(sample_text("Title", Language::English), "Document Title");
        assert_eq!(
            sample_text("Heading 1", Language::English),
            "Heading 1 Example"
        );
        assert_eq!(
            sample_text("Normal paragraph", Language::English),
            "This is a normal paragraph. You can modify this style in Word to change the default body text formatting."
        );
    }

    #[test]
    fn test_sample_text_thai() {
        assert_eq!(
            sample_text("Title", Language::Thai),
            "ชื่อเอกสาร (Document Title)"
        );
        assert_eq!(
            sample_text("Heading 1", Language::Thai),
            "หัวข้อ 1 (Heading 1)"
        );
        assert!(sample_text("Normal paragraph", Language::Thai).contains("ย่อหน้า"));
    }

    #[test]
    fn test_sample_text_unknown_style() {
        assert_eq!(sample_text("Unknown", Language::English), "Sample text");
        assert_eq!(sample_text("Unknown", Language::Thai), "Sample text");
    }
}
