//! md2docx - Markdown to DOCX converter with Thai/English support
//!
//! # Quick Start
//!
//! ```rust,no_run
//! use md2docx::{Document, Language};
//!
//! let doc = Document::new()
//!     .add_heading(1, "Hello World")
//!     .add_paragraph("This is a test.");
//!
//! doc.write_to_file("output.docx").unwrap();
//! ```

pub mod config;
pub mod discovery;
pub mod docx;
pub mod error;
pub mod i18n;
pub mod parser;

pub use docx::ooxml::{FooterConfig, HeaderConfig, HeaderFooterField};
pub use docx::toc::TocConfig;
pub use docx::DocumentConfig;
pub use parser::{IncludeConfig, IncludeResolver, ParsedDocument};

#[cfg(feature = "mermaid-cli")]
pub mod mermaid;

#[cfg(feature = "git")]
pub mod diff;

#[cfg(feature = "wasm")]
pub mod wasm;

pub use docx::ooxml::{FootnotesXml, Language, Paragraph, Run};
pub use error::{Error, Result};

use docx::ooxml::{
    generate_numbering_xml_with_context, ContentTypes, DocumentXml, Relationships, StylesDocument,
};
use docx::{build_document, Packager};
use parser::parse_markdown_with_frontmatter;
use std::io::Cursor;

/// High-level Document builder for creating DOCX files
#[derive(Debug)]
pub struct Document {
    /// Document content
    doc_xml: DocumentXml,
    /// Language for styles/fonts
    lang: Language,
}

impl Document {
    /// Create a new empty document
    pub fn new() -> Self {
        Self {
            doc_xml: DocumentXml::new(),
            lang: Language::English,
        }
    }

    /// Create a new document with specific language
    pub fn with_language(lang: Language) -> Self {
        Self {
            doc_xml: DocumentXml::new(),
            lang,
        }
    }

    /// Set the document language
    pub fn language(mut self, lang: Language) -> Self {
        self.lang = lang;
        self
    }

    /// Add a heading (level 1-4)
    pub fn add_heading(mut self, level: u8, text: &str) -> Self {
        let style_id = match level {
            1 => "Heading1",
            2 => "Heading2",
            3 => "Heading3",
            _ => "Heading4",
        };
        let p = Paragraph::with_style(style_id)
            .add_text(text)
            .spacing(0, 0)
            .line_spacing(240, "auto");
        self.doc_xml.add_paragraph(p);
        self
    }

    /// Add a paragraph with optional style
    pub fn add_paragraph(mut self, text: &str) -> Self {
        let p = Paragraph::with_style("Normal")
            .add_text(text)
            .spacing(0, 0)
            .line_spacing(240, "auto");
        self.doc_xml.add_paragraph(p);
        self
    }

    /// Add a styled paragraph
    pub fn add_styled_paragraph(mut self, style: &str, text: &str) -> Self {
        let p = Paragraph::with_style(style)
            .add_text(text)
            .spacing(0, 0)
            .line_spacing(240, "auto");
        self.doc_xml.add_paragraph(p);
        self
    }

    /// Add a paragraph with a Run (for fine-grained control)
    pub fn add_paragraph_with_runs(mut self, style: &str, runs: Vec<Run>) -> Self {
        let mut p = Paragraph::with_style(style)
            .spacing(0, 0)
            .line_spacing(240, "auto");
        for run in runs {
            p = p.add_run(run);
        }
        self.doc_xml.add_paragraph(p);
        self
    }

    /// Add a code block
    pub fn add_code_block(mut self, code: &str) -> Self {
        // Split by lines and add each as a Code paragraph
        for line in code.lines() {
            let p = Paragraph::with_style("Code")
                .add_text(line)
                .spacing(0, 0)
                .line_spacing(240, "auto");
            self.doc_xml.add_paragraph(p);
        }
        self
    }

    /// Add a blockquote
    pub fn add_quote(mut self, text: &str) -> Self {
        let p = Paragraph::with_style("Quote").add_text(text);
        self.doc_xml.add_paragraph(p);
        self
    }

    /// Add raw paragraph (for advanced use)
    pub fn add_raw_paragraph(mut self, paragraph: Paragraph) -> Self {
        self.doc_xml.add_paragraph(paragraph);
        self
    }

    /// Build the document and return bytes
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Create components
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();
        let styles = StylesDocument::new(self.lang);

        // Package
        packager.package(
            &self.doc_xml,
            &styles,
            &content_types,
            &rels,
            &doc_rels,
            self.lang,
        )?;

        let cursor = packager.finish()?;
        Ok(cursor.into_inner())
    }

    /// Write document to a file (only available when not targeting WASM)
    #[cfg(not(target_arch = "wasm32"))]
    pub fn write_to_file(&self, path: impl AsRef<std::path::Path>) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

impl Default for Document {
    fn default() -> Self {
        Self::new()
    }
}

/// Convert markdown string to DOCX bytes
///
/// This is a convenience function that parses markdown and generates
/// a complete DOCX file in one step.
///
/// # Arguments
/// * `markdown` - The markdown content to convert
///
/// # Returns
/// A `Result` containing the DOCX file as bytes
///
/// # Example
/// ```rust,no_run
/// use md2docx::markdown_to_docx;
///
/// let md = "# Hello World\n\nThis is **bold** text.";
/// let docx_bytes = markdown_to_docx(md).unwrap();
/// std::fs::write("output.docx", docx_bytes).unwrap();
/// ```
pub fn markdown_to_docx(markdown: &str) -> Result<Vec<u8>> {
    markdown_to_docx_with_config(markdown, Language::English, &DocumentConfig::default())
}

/// Convert markdown string to DOCX bytes with custom configuration
///
/// This function allows you to specify language and document configuration
/// such as TOC, headers, footers, etc.
///
/// # Arguments
/// * `markdown` - The markdown content to convert
/// * `lang` - The language for fonts and styles (English or Thai)
/// * `config` - Document configuration (TOC, headers, footers, etc.)
///
/// # Returns
/// A `Result` containing the DOCX file as bytes
///
/// # Example
/// ```rust,no_run
/// use md2docx::{markdown_to_docx_with_config, DocumentConfig, Language};
///
/// let md = "# Hello World\n\nThis is **bold** text.";
/// let config = DocumentConfig {
///     title: "My Document".to_string(),
///     ..Default::default()
/// };
/// let docx_bytes = markdown_to_docx_with_config(md, Language::English, &config).unwrap();
/// std::fs::write("output.docx", docx_bytes).unwrap();
/// ```
pub fn markdown_to_docx_with_config(
    markdown: &str,
    lang: Language,
    config: &DocumentConfig,
) -> Result<Vec<u8>> {
    let parsed = parse_markdown_with_frontmatter(markdown);

    let mut build_result = build_document(&parsed, lang, config);

    let buffer = Cursor::new(Vec::new());
    let mut packager = Packager::new(buffer);

    let mut content_types = ContentTypes::new();
    let rels = Relationships::root_rels();
    let mut doc_rels = Relationships::document_rels();
    let styles = StylesDocument::new(lang);

    // Process images
    for image in &build_result.images.images {
        // 1. Add content type
        let ext = std::path::Path::new(&image.filename)
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("png");

        let content_type = match ext.to_lowercase().as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "svg" => "image/svg+xml",
            _ => "application/octet-stream",
        };
        content_types.add_image_extension(ext, content_type);

        // 2. Add relationship
        // Use add_image_with_id to ensure the relationship ID matches what's in the document.xml
        // The ImageContext generates specific IDs (rId4, rId5...) which we must preserve.
        doc_rels.add_image_with_id(&image.rel_id, &image.filename);

        // Sanity check (optional but good for debugging)
        if image.rel_id.starts_with("rId") {
            // We trust the image.rel_id
        }

        // 3. Add image file (CLI only)
        #[cfg(not(target_arch = "wasm32"))]
        {
            // Try to read file. If fails, we skip adding the file but keep the ref.
            // Word will show a placeholder with "image not found".
            if let Ok(data) = std::fs::read(&image.src) {
                packager.add_image(&image.filename, &data)?;
            }
            // If file read fails, silent failure or placeholder is fine for MVP
        }
    }

    // Always add footnotes.xml (settings.xml references footnote IDs -1 and 0)
    // Even with no user footnotes, the separators must exist
    content_types.add_footnotes();
    doc_rels.add_footnotes();
    let footnotes_xml = build_result.footnotes.to_xml()?;
    packager.add_footnotes(&footnotes_xml)?;

    // Always add endnotes.xml (settings.xml references endnote IDs -1 and 0)
    use crate::docx::ooxml::EndnotesXml;
    content_types.add_endnotes();
    doc_rels.add_endnotes();
    let endnotes = EndnotesXml::new();
    let endnotes_xml = endnotes.to_xml()?;
    packager.add_endnotes(&endnotes_xml)?;

    // Process hyperlinks
    for link in &build_result.hyperlinks.hyperlinks {
        doc_rels.add_hyperlink_with_id(&link.rel_id, &link.url);
    }

    // Always add numbering.xml for list support
    content_types.add_numbering();
    doc_rels.add_numbering();
    let numbering_xml = generate_numbering_xml_with_context(&build_result.numbering)?;
    packager.add_numbering(&numbering_xml)?;

    // Process headers and capture returned relationship IDs
    let mut header_rel_ids: Vec<(u32, String)> = Vec::new();
    for (header_num, _) in &build_result.headers {
        content_types.add_header(*header_num);
        let rel_id = doc_rels.add_header(*header_num);
        header_rel_ids.push((*header_num, rel_id));
    }

    // Process footers and capture returned relationship IDs
    let mut footer_rel_ids: Vec<(u32, String)> = Vec::new();
    for (footer_num, _) in &build_result.footers {
        content_types.add_footer(*footer_num);
        let rel_id = doc_rels.add_footer(*footer_num);
        footer_rel_ids.push((*footer_num, rel_id));
    }

    // Update header/footer refs with actual relationship IDs from doc_rels
    // Header 1 is default, header 2 is first page
    for (num, rel_id) in &header_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_header_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_header_id = Some(rel_id.clone());
            // Also store as empty header ID for suppression
            build_result.document.empty_header_id = Some(rel_id.clone());
        }
    }

    for (num, rel_id) in &footer_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_footer_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_footer_id = Some(rel_id.clone());
            // Also store as empty footer ID for suppression
            build_result.document.empty_footer_id = Some(rel_id.clone());
        }
    }

    packager.package(
        &build_result.document,
        &styles,
        &content_types,
        &rels,
        &doc_rels,
        lang,
    )?;

    // Add headers to the archive
    for (header_num, header_bytes) in &build_result.headers {
        packager.add_header(*header_num, header_bytes)?;
    }

    // Add footers to the archive
    for (footer_num, footer_bytes) in &build_result.footers {
        packager.add_footer(*footer_num, footer_bytes)?;
    }

    let cursor = packager.finish()?;
    Ok(cursor.into_inner())
}

/// Resolve include directives in a parsed document
///
/// This function processes `{!include:...}` and `{!code:...}` directives
/// by loading external files and converting them to markdown blocks.
///
/// # Arguments
/// * `doc` - The parsed document to resolve includes in (modified in place)
/// * `config` - Configuration for include resolution (base paths, max depth, etc.)
///
/// # Returns
/// A `Result` indicating success or failure
///
/// # Example
/// ```rust,no_run
/// use md2docx::{resolve_includes, IncludeConfig, ParsedDocument};
/// use std::path::PathBuf;
///
/// // First parse your markdown (using your own parser or md2docx's internal parser)
/// // let mut parsed = ParsedDocument { /* ... */ };
///
/// let config = IncludeConfig {
///     base_path: PathBuf::from("./docs"),
///     source_root: PathBuf::from("./src"),
///     max_depth: 10,
/// };
///
/// // resolve_includes(&mut parsed, &config).unwrap();
/// // Now parsed.blocks has the included content expanded
/// ```
///
/// # Note
/// This function requires file system access and is not available in WASM builds.
#[cfg(not(target_arch = "wasm32"))]
pub fn resolve_includes(doc: &mut ParsedDocument, config: &IncludeConfig) -> Result<()> {
    let mut resolver = IncludeResolver::new(config.clone());
    doc.blocks = resolver.resolve_blocks(std::mem::take(&mut doc.blocks))?;
    Ok(())
}

/// Convert markdown to DOCX with include resolution
///
/// This is a convenience function that parses markdown, resolves include directives,
/// and generates a complete DOCX file in one step.
///
/// # Arguments
/// * `markdown` - The markdown content to convert
/// * `include_config` - Configuration for include resolution (optional)
///
/// # Returns
/// A `Result` containing the DOCX file as bytes
///
/// # Example
/// ```rust,no_run
/// use md2docx::{markdown_to_docx_with_includes, IncludeConfig};
/// use std::path::PathBuf;
///
/// let md = "# Introduction\n\n{!include:section.md}";
///
/// let include_config = IncludeConfig {
///     base_path: PathBuf::from("./docs"),
///     source_root: PathBuf::from("./src"),
///     max_depth: 10,
/// };
///
/// let docx_bytes = markdown_to_docx_with_includes(md, &include_config).unwrap();
/// std::fs::write("output.docx", docx_bytes).unwrap();
/// ```
///
/// # Note
/// This function requires file system access and is not available in WASM builds.
#[cfg(not(target_arch = "wasm32"))]
pub fn markdown_to_docx_with_includes(
    markdown: &str,
    include_config: &IncludeConfig,
) -> Result<Vec<u8>> {
    let mut parsed = parse_markdown_with_frontmatter(markdown);

    // Resolve includes
    resolve_includes(&mut parsed, include_config)?;

    // Determine language from frontmatter, default to English
    let lang = if let Some(ref fm) = parsed.frontmatter {
        if let Some(ref l) = fm.language {
            match l.to_lowercase().as_str() {
                "th" | "thai" => Language::Thai,
                _ => Language::English,
            }
        } else {
            Language::English
        }
    } else {
        Language::English
    };

    let mut build_result = build_document(&parsed, lang, &DocumentConfig::default());

    let buffer = Cursor::new(Vec::new());
    let mut packager = Packager::new(buffer);

    let mut content_types = ContentTypes::new();
    let rels = Relationships::root_rels();
    let mut doc_rels = Relationships::document_rels();
    let styles = StylesDocument::new(lang);

    // Process images
    for image in &build_result.images.images {
        let ext = std::path::Path::new(&image.filename)
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("png");

        let content_type = match ext.to_lowercase().as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "svg" => "image/svg+xml",
            _ => "application/octet-stream",
        };
        content_types.add_image_extension(ext, content_type);

        doc_rels.add_image_with_id(&image.rel_id, &image.filename);

        if image.rel_id.starts_with("rId") {
            // Valid relationship ID
        }

        #[cfg(not(target_arch = "wasm32"))]
        {
            if let Ok(data) = std::fs::read(&image.src) {
                packager.add_image(&image.filename, &data)?;
            }
        }
    }

    // Always add footnotes.xml (settings.xml references footnote IDs -1 and 0)
    content_types.add_footnotes();
    doc_rels.add_footnotes();
    let footnotes_xml = build_result.footnotes.to_xml()?;
    packager.add_footnotes(&footnotes_xml)?;

    // Always add endnotes.xml (settings.xml references endnote IDs -1 and 0)
    use crate::docx::ooxml::EndnotesXml;
    content_types.add_endnotes();
    doc_rels.add_endnotes();
    let endnotes = EndnotesXml::new();
    let endnotes_xml = endnotes.to_xml()?;
    packager.add_endnotes(&endnotes_xml)?;

    // Process hyperlinks
    for link in &build_result.hyperlinks.hyperlinks {
        doc_rels.add_hyperlink_with_id(&link.rel_id, &link.url);
    }

    // Always add numbering.xml for list support
    content_types.add_numbering();
    doc_rels.add_numbering();
    let numbering_xml = generate_numbering_xml_with_context(&build_result.numbering)?;
    packager.add_numbering(&numbering_xml)?;

    // Process headers and capture returned relationship IDs
    let mut header_rel_ids: Vec<(u32, String)> = Vec::new();
    for (header_num, _) in &build_result.headers {
        content_types.add_header(*header_num);
        let rel_id = doc_rels.add_header(*header_num);
        header_rel_ids.push((*header_num, rel_id));
    }

    // Process footers and capture returned relationship IDs
    let mut footer_rel_ids: Vec<(u32, String)> = Vec::new();
    for (footer_num, _) in &build_result.footers {
        content_types.add_footer(*footer_num);
        let rel_id = doc_rels.add_footer(*footer_num);
        footer_rel_ids.push((*footer_num, rel_id));
    }

    // Update header/footer refs with actual relationship IDs from doc_rels
    // Header 1 is default, header 2 is first page
    for (num, rel_id) in &header_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_header_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_header_id = Some(rel_id.clone());
            // Also store as empty header ID for suppression
            build_result.document.empty_header_id = Some(rel_id.clone());
        }
    }

    for (num, rel_id) in &footer_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_footer_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_footer_id = Some(rel_id.clone());
            // Also store as empty footer ID for suppression
            build_result.document.empty_footer_id = Some(rel_id.clone());
        }
    }

    packager.package(
        &build_result.document,
        &styles,
        &content_types,
        &rels,
        &doc_rels,
        lang,
    )?;

    // Add headers to the archive
    for (header_num, header_bytes) in &build_result.headers {
        packager.add_header(*header_num, header_bytes)?;
    }

    // Add footers to the archive
    for (footer_num, footer_bytes) in &build_result.footers {
        packager.add_footer(*footer_num, footer_bytes)?;
    }

    let cursor = packager.finish()?;
    Ok(cursor.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx::ooxml::DocElement;
    use parser::{Block, Inline};
    use std::path::PathBuf;

    /// Helper function to extract paragraphs from document elements
    fn get_paragraphs(doc: &Document) -> Vec<&Paragraph> {
        doc.doc_xml
            .elements
            .iter()
            .filter_map(|e| match e {
                DocElement::Paragraph(p) => Some(p.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_document_default() {
        let doc = Document::new();
        assert_eq!(doc.lang, Language::English);
    }

    #[test]
    fn test_document_with_language() {
        let doc = Document::with_language(Language::Thai);
        assert_eq!(doc.lang, Language::Thai);
    }

    #[test]
    fn test_document_language_setter() {
        let doc = Document::new().language(Language::Thai);
        assert_eq!(doc.lang, Language::Thai);
    }

    #[test]
    fn test_add_heading() {
        let doc = Document::new()
            .add_heading(1, "Title")
            .add_heading(2, "Subtitle");

        // Should have 2 paragraphs
        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_id, Some("Heading1".to_string()));
        assert_eq!(paragraphs[1].style_id, Some("Heading2".to_string()));
    }

    #[test]
    fn test_add_paragraph() {
        let doc = Document::new()
            .add_paragraph("First paragraph")
            .add_paragraph("Second paragraph");

        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_id, Some("Normal".to_string()));
        assert_eq!(paragraphs[1].style_id, Some("Normal".to_string()));
    }

    #[test]
    fn test_add_styled_paragraph() {
        let doc = Document::new().add_styled_paragraph("Quote", "This is a quote");

        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].style_id, Some("Quote".to_string()));
    }

    #[test]
    fn test_add_paragraph_with_runs() {
        let runs = vec![
            Run::new("Normal "),
            Run::new("Bold").bold(),
            Run::new(" Italic").italic(),
        ];

        let doc = Document::new().add_paragraph_with_runs("Normal", runs);

        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].children.len(), 3);
    }

    #[test]
    fn test_add_code_block() {
        let code = "fn main() {\n    println!(\"Hello\");\n}";
        let doc = Document::new().add_code_block(code);

        // Should have 3 lines
        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 3);
        assert_eq!(paragraphs[0].style_id, Some("Code".to_string()));
        assert_eq!(paragraphs[1].style_id, Some("Code".to_string()));
        assert_eq!(paragraphs[2].style_id, Some("Code".to_string()));
    }

    #[test]
    fn test_add_quote() {
        let doc = Document::new().add_quote("This is a quote");

        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].style_id, Some("Quote".to_string()));
    }

    #[test]
    fn test_add_raw_paragraph() {
        let p = Paragraph::with_style("Heading1").add_text("Custom");
        let doc = Document::new().add_raw_paragraph(p);

        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].style_id, Some("Heading1".to_string()));
    }

    #[test]
    fn test_to_bytes() {
        let doc = Document::new()
            .add_heading(1, "Test Document")
            .add_paragraph("This is a test.");

        let bytes = doc.to_bytes().unwrap();

        // Should have some data
        assert!(!bytes.is_empty());

        // Should be a valid ZIP (starts with PK magic bytes)
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_to_bytes_thai() {
        let doc = Document::with_language(Language::Thai)
            .add_heading(1, "เอกสารทดสอบ")
            .add_paragraph("นี่คือการทดสอบภาษาไทย");

        let bytes = doc.to_bytes().unwrap();

        assert!(!bytes.is_empty());
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_complex_document() {
        let doc = Document::new()
            .add_heading(1, "Chapter 1")
            .add_heading(2, "Section 1.1")
            .add_paragraph("This is the first paragraph.")
            .add_paragraph("This is the second paragraph.")
            .add_quote("This is a quote.")
            .add_code_block("fn main() {\n    println!(\"Hello\");\n}")
            .add_heading(3, "Subsection")
            .add_paragraph("More content.");

        // Code block has 3 lines: "fn main() {", "    println!(\"Hello\");", "}"
        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 10);

        let bytes = doc.to_bytes().unwrap();
        assert!(!bytes.is_empty());
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_builder_pattern_chain() {
        let doc = Document::new()
            .add_heading(1, "Title")
            .add_paragraph("Content")
            .add_quote("Quote")
            .add_code_block("code");

        assert_eq!(get_paragraphs(&doc).len(), 4);
    }

    #[test]
    fn test_heading_level_4() {
        let doc = Document::new()
            .add_heading(4, "Level 4")
            .add_heading(5, "Level 5"); // Should also use Heading4

        let paragraphs = get_paragraphs(&doc);
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_id, Some("Heading4".to_string()));
        assert_eq!(paragraphs[1].style_id, Some("Heading4".to_string()));
    }

    #[test]
    fn test_empty_document() {
        let doc = Document::new();
        let bytes = doc.to_bytes().unwrap();

        // Even empty document should produce valid ZIP
        assert!(!bytes.is_empty());
        assert_eq!(&bytes[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_document_with_mixed_content() {
        let runs = vec![
            Run::new("Normal text "),
            Run::new("bold text").bold(),
            Run::new(" and "),
            Run::new("italic text").italic(),
        ];

        let doc = Document::new()
            .add_heading(1, "Mixed Content")
            .add_paragraph_with_runs("Normal", runs)
            .add_code_block("line 1\nline 2")
            .add_quote("A quote");

        assert_eq!(get_paragraphs(&doc).len(), 5);
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_write_to_file() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_output.docx");

        let doc = Document::new()
            .add_heading(1, "Test")
            .add_paragraph("Content");

        doc.write_to_file(&file_path).unwrap();

        // Verify file exists
        assert!(file_path.exists());

        // Verify it's a valid ZIP
        let contents = fs::read(&file_path).unwrap();
        assert_eq!(&contents[0..4], b"PK\x03\x04");

        // Cleanup
        fs::remove_file(&file_path).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_write_to_file_thai() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let file_path = temp_dir.join("test_thai.docx");

        let doc = Document::with_language(Language::Thai)
            .add_heading(1, "ทดสอบ")
            .add_paragraph("ภาษาไทย");

        doc.write_to_file(&file_path).unwrap();

        assert!(file_path.exists());

        let contents = fs::read(&file_path).unwrap();
        assert_eq!(&contents[0..4], b"PK\x03\x04");

        fs::remove_file(&file_path).unwrap();
    }

    #[test]
    fn test_markdown_to_docx_with_header_footer() {
        use crate::docx::ooxml::HeaderFooterField;

        let md = "# Test Document\n\nThis is a test.";

        // Create config with header and footer
        let config = DocumentConfig {
            title: "Test Title".to_string(),
            header: HeaderConfig {
                left: vec![HeaderFooterField::DocumentTitle],
                center: vec![],
                right: vec![HeaderFooterField::ChapterName],
            },
            footer: FooterConfig {
                left: vec![],
                center: vec![HeaderFooterField::PageNumber],
                right: vec![],
            },
            different_first_page: false,
            ..Default::default()
        };

        let parsed = parse_markdown_with_frontmatter(md);
        let build_result = build_document(&parsed, Language::English, &config);

        // Verify headers and footers were generated
        assert!(
            !build_result.headers.is_empty(),
            "Should have at least one header"
        );
        assert!(
            !build_result.footers.is_empty(),
            "Should have at least one footer"
        );

        // Package the document
        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        let mut content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let mut doc_rels = Relationships::document_rels();
        let styles = StylesDocument::new(Language::English);

        // Add header/footer content types (header1, footer1)
        content_types.add_header(1);
        content_types.add_footer(1);

        // Add header/footer relationships (returns the relationship IDs)
        let header_rel_id = doc_rels.add_header(1);
        let footer_rel_id = doc_rels.add_footer(1);

        packager
            .package(
                &build_result.document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add headers and footers
        for (header_num, header_bytes) in &build_result.headers {
            packager.add_header(*header_num, header_bytes).unwrap();
        }

        for (footer_num, footer_bytes) in &build_result.footers {
            packager.add_footer(*footer_num, footer_bytes).unwrap();
        }

        let cursor = packager.finish().unwrap();
        let zip_data = cursor.into_inner();

        // Verify we got a valid ZIP
        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");

        // Verify relationship IDs were generated
        assert!(header_rel_id.starts_with("rId"));
        assert!(footer_rel_id.starts_with("rId"));
    }

    #[test]
    fn test_document_config_exports() {
        // Verify that DocumentConfig, HeaderConfig, FooterConfig, and HeaderFooterField are exported
        let _config = DocumentConfig::default();
        let _header = HeaderConfig::default();
        let _footer = FooterConfig::default();
        let _field = HeaderFooterField::Text("test".to_string());
        let _field2 = HeaderFooterField::PageNumber;
        let _field3 = HeaderFooterField::ChapterName;
        let _field4 = HeaderFooterField::DocumentTitle;
        let _field5 = HeaderFooterField::TotalPages;
    }

    #[test]
    fn test_markdown_to_docx_includes_headers_footers() {
        // Test that markdown_to_docx properly packages headers and footers
        let md = "# Test Document\n\nThis is a test.";

        // The default DocumentConfig has headers and footers
        let docx_bytes = markdown_to_docx(md).unwrap();

        // Verify we got a valid ZIP
        assert!(!docx_bytes.is_empty());
        assert_eq!(&docx_bytes[0..4], b"PK\x03\x04");

        // The default config should generate headers and footers
        // We can verify this by checking that the build result includes them
        let parsed = parse_markdown_with_frontmatter(md);
        let build_result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Default config has headers and footers
        assert!(
            !build_result.headers.is_empty(),
            "Default config should have headers"
        );
        assert!(
            !build_result.footers.is_empty(),
            "Default config should have footers"
        );
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_resolve_includes_function() {
        use std::fs;
        use std::io::Write;

        let temp_dir = std::env::temp_dir();
        let test_dir = temp_dir.join("md2docx_test_includes");
        fs::create_dir_all(&test_dir).unwrap();

        // Create a test file to include
        let included_file = test_dir.join("included.md");
        let mut file = fs::File::create(&included_file).unwrap();
        file.write_all(b"This is included content").unwrap();

        // Create a test code file
        let code_file = test_dir.join("code.rs");
        let mut file = fs::File::create(&code_file).unwrap();
        file.write_all(b"fn main() {\n    println!(\"Hello\");\n}")
            .unwrap();

        let mut doc = ParsedDocument {
            frontmatter: None,
            blocks: vec![
                Block::Paragraph(vec![Inline::Text("Hello".to_string())]),
                Block::Include {
                    path: "included.md".to_string(),
                    resolved: None,
                },
                Block::CodeInclude {
                    path: "code.rs".to_string(),
                    start_line: None,
                    end_line: None,
                    lang: None,
                },
            ],
            footnotes: std::collections::HashMap::new(),
        };

        let config = IncludeConfig {
            base_path: test_dir.clone(),
            source_root: test_dir.clone(),
            max_depth: 10,
        };

        let result = resolve_includes(&mut doc, &config);
        assert!(
            result.is_ok(),
            "resolve_includes should succeed: {:?}",
            result
        );

        // Verify includes were resolved
        assert_eq!(doc.blocks.len(), 3);
        assert!(matches!(doc.blocks[0], Block::Paragraph(_)));
        assert!(matches!(doc.blocks[1], Block::Paragraph(_))); // Included content
        assert!(matches!(doc.blocks[2], Block::CodeBlock { .. })); // Code include

        // Cleanup
        fs::remove_dir_all(test_dir).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_resolve_includes_nonexistent_file() {
        let mut doc = ParsedDocument {
            frontmatter: None,
            blocks: vec![
                Block::Paragraph(vec![Inline::Text("Hello".to_string())]),
                Block::CodeInclude {
                    path: "nonexistent.rs".to_string(),
                    start_line: None,
                    end_line: None,
                    lang: None,
                },
            ],
            footnotes: std::collections::HashMap::new(),
        };

        let config = IncludeConfig::default();
        // This should fail because the file doesn't exist
        let result = resolve_includes(&mut doc, &config);
        assert!(result.is_err(), "Should fail for nonexistent file");
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_include_config_exports() {
        // Verify that IncludeConfig is exported and usable
        let config = IncludeConfig {
            base_path: PathBuf::from("."),
            source_root: PathBuf::from("./src"),
            max_depth: 10,
        };
        assert_eq!(config.max_depth, 10);

        let default_config = IncludeConfig::default();
        assert_eq!(default_config.max_depth, 10);
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_parsed_document_exports() {
        // Verify that ParsedDocument is exported and usable
        let doc = ParsedDocument {
            frontmatter: None,
            blocks: vec![Block::Paragraph(vec![Inline::Text("Test".to_string())])],
            footnotes: std::collections::HashMap::new(),
        };
        assert_eq!(doc.blocks.len(), 1);
    }
}
