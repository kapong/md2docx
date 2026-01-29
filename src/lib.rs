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
pub mod template;

#[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
pub mod project;

pub use docx::ooxml::{FooterConfig, HeaderConfig, HeaderFooterField};
pub use docx::toc::TocConfig;
pub use docx::{DocumentConfig, DocumentMeta};
pub use parser::{IncludeConfig, IncludeResolver, ParsedDocument};
pub use template::{PlaceholderContext, TemplateDir, TemplateSet};

// Re-export template extraction types for use in examples
pub use template::extract::{CoverTemplate, HeaderFooterTemplate, ImageTemplate, TableTemplate};

// Re-export helper function for finding image paths
pub use template::extract::cover::find_image_path_from_rel_id;

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
        let styles = StylesDocument::new(self.lang, None);

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
    markdown_to_docx_with_templates(markdown, lang, config, None, &PlaceholderContext::default())
}

/// Convert markdown to DOCX with template support
///
/// This function extends `markdown_to_docx_with_config` by adding support for
/// template directories containing cover.docx, table.docx, image.docx, and
/// header-footer.docx files.
///
/// # Arguments
/// * `markdown` - The markdown content to convert
/// * `lang` - The language for fonts and styles
/// * `config` - Document configuration
/// * `templates` - Optional template set loaded from template directory
/// * `placeholder_ctx` - Context for placeholder replacement in templates
///
/// # Returns
/// A `Result` containing the DOCX file as bytes
///
/// # Example
/// ```rust,no_run
/// use md2docx::{
///     markdown_to_docx_with_templates, DocumentConfig, Language,
///     PlaceholderContext, TemplateSet
/// };
///
/// let md = "# Hello World\n\nThis is content.";
/// let config = DocumentConfig::default();
/// let templates = TemplateSet::default(); // Load from TemplateDir
/// let ctx = PlaceholderContext::default();
///
/// let docx_bytes = markdown_to_docx_with_templates(
///     md, Language::English, &config, Some(&templates), &ctx
/// ).unwrap();
/// ```
/// Convert markdown to DOCX with template support
///
/// This function extends `markdown_to_docx_with_config` by adding support for
/// template directories containing cover.docx, table.docx, image.docx, and
/// header-footer.docx files.
///
/// # Arguments
/// * `markdown` - The markdown content to convert
/// * `lang` - The language for fonts and styles
/// * `config` - Document configuration
/// * `templates` - Optional template set loaded from template directory
/// * `placeholder_ctx` - Context for placeholder replacement in templates
///
/// # Returns
/// A `Result` containing the DOCX file as bytes
///
/// # Example
/// ```rust,no_run
/// use md2docx::{
///     markdown_to_docx_with_templates, DocumentConfig, Language,
///     PlaceholderContext, TemplateSet
/// };
///
/// let md = "# Hello World\n\nThis is content.";
/// let config = DocumentConfig::default();
/// let templates = TemplateSet::default(); // Load from TemplateDir
/// let ctx = PlaceholderContext::default();
///
/// let docx_bytes = markdown_to_docx_with_templates(
///     md, Language::English, &config, Some(&templates), &ctx
/// ).unwrap();
/// ```
pub fn markdown_to_docx_with_templates(
    markdown: &str,
    lang: Language,
    doc_config: &DocumentConfig,
    templates: Option<&crate::template::TemplateSet>,
    placeholder_ctx: &crate::template::PlaceholderContext,
) -> Result<Vec<u8>> {
    let parsed = parse_markdown_with_frontmatter(markdown);

    let mut rel_manager = crate::docx::RelIdManager::new();
    let table_template = templates.and_then(|t| t.table.as_ref());
    let image_template = templates.and_then(|t| t.image.as_ref());
    let mut build_result = build_document(
        &parsed,
        lang,
        doc_config,
        &mut rel_manager,
        table_template,
        image_template,
    );

    // Apply templates if provided
    if let Some(template_set) = templates {
        // Apply cover template
        if let Some(cover) = &template_set.cover {
            apply_cover_template(
                &mut build_result,
                cover,
                placeholder_ctx,
                lang,
                &mut rel_manager,
                table_template,
                image_template,
                doc_config,
            )?;
        }
    }

    // Insert TOC if enabled
    if let Some(toc_builder) = build_result.toc_builder.take() {
        if doc_config.toc.enabled && !toc_builder.is_empty() {
            let toc_elements = toc_builder.generate_toc(&doc_config.toc);

            // Determine insertion position
            let mut has_cover = false;
            if let Some(t) = templates {
                if t.cover.is_some() {
                    has_cover = true;
                }
            }

            if has_cover {
                // Find the section break after cover (should be at index 1)
                if build_result.document.elements.len() > 1 {
                    if let crate::docx::ooxml::DocElement::Paragraph(p) =
                        &mut build_result.document.elements[1]
                    {
                        if p.is_section_break() {
                            // Change section break to page break to keep TOC in same section as cover
                            *p = Box::new(crate::docx::ooxml::Paragraph::new().page_break());
                        }
                    }
                }

                // Insert TOC elements after the page break
                for (i, elem) in toc_elements.into_iter().enumerate() {
                    build_result.document.elements.insert(2 + i, elem);
                }
            } else {
                // No cover, insert at the beginning
                for (i, elem) in toc_elements.into_iter().enumerate() {
                    build_result.document.elements.insert(i, elem);
                }
            }
        }
    }

    // Ensure Chapter 1 starts at page 1
    // Find the first Heading 1 (start of Chapter 1)
    let mut chapter1_index = None;
    for (i, elem) in build_result.document.elements.iter().enumerate() {
        if let crate::docx::ooxml::DocElement::Paragraph(p) = elem {
            if p.style_id.as_deref() == Some("Heading1") {
                chapter1_index = Some(i);
                break;
            }
        }
    }

    if let Some(idx) = chapter1_index {
        // We found Chapter 1. Now we need to set page numbering restart on the section properties
        // that apply to Chapter 1.
        // In DOCX, section properties are defined at the END of the section (in a section break),
        // or at the end of the document (w:sectPr) for the final section.

        // Look for the next section break after Chapter 1 start
        let mut found_next_break = false;
        for i in (idx + 1)..build_result.document.elements.len() {
            if let crate::docx::ooxml::DocElement::Paragraph(p) =
                &mut build_result.document.elements[i]
            {
                if p.is_section_break() {
                    // Found the section break that ends Chapter 1 (and defines its properties)
                    p.page_num_start = Some(1);
                    found_next_break = true;
                    break;
                }
            }
        }

        // If no section break found after Chapter 1, it means Chapter 1 is the last section.
        // Its properties are defined in the document's final sectPr.
        if !found_next_break {
            build_result.document.page_num_start = Some(1);
        }

        // Also check if there's a section break *before* Chapter 1 (e.g. from TOC).
        // That section break defines properties for the TOC section.
        // We should ensure that section break DOES NOT restart numbering (or restarts at something else if needed),
        // but typically TOC uses Roman numerals or standard numbering.
        // The previous code was incorrectly setting page_num_start on the TOC section break.
    }

    // Note: Table and image templates would be applied during block processing
    // This requires modifying the builder to use template styles
    // For now, we just load and extract the templates

    let buffer = Cursor::new(Vec::new());
    let mut packager = Packager::new(buffer);

    let mut content_types = ContentTypes::new();
    let rels = Relationships::root_rels();
    let mut doc_rels = Relationships::document_rels();
    let styles = StylesDocument::new(lang, doc_config.fonts.clone());

    // Process images from build_result (includes cover template images and markdown images)
    // Header/footer images are handled separately with header_ prefix
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

        #[cfg(not(target_arch = "wasm32"))]
        {
            if let Some(ref data) = image.data {
                packager.add_image(&image.filename, data)?;
            } else if let Ok(data) = std::fs::read(&image.src) {
                packager.add_image(&image.filename, &data)?;
            }
        }
    }

    // Add footnotes
    content_types.add_footnotes();
    let footnotes_rel_id = rel_manager.next_id();
    doc_rels.add_footnotes_with_id(&footnotes_rel_id);
    let footnotes_xml = build_result.footnotes.to_xml()?;
    packager.add_footnotes(&footnotes_xml)?;

    // Add endnotes
    use crate::docx::ooxml::EndnotesXml;
    content_types.add_endnotes();
    let endnotes_rel_id = rel_manager.next_id();
    doc_rels.add_endnotes_with_id(&endnotes_rel_id);
    let endnotes = EndnotesXml::new();
    let endnotes_xml = endnotes.to_xml()?;
    packager.add_endnotes(&endnotes_xml)?;

    // Process hyperlinks
    for link in &build_result.hyperlinks.hyperlinks {
        doc_rels.add_hyperlink_with_id(&link.rel_id, &link.url);
    }

    // Add numbering
    content_types.add_numbering();
    let numbering_rel_id = rel_manager.next_id();
    doc_rels.add_numbering_with_id(&numbering_rel_id);
    let numbering_xml = generate_numbering_xml_with_context(&build_result.numbering)?;
    packager.add_numbering(&numbering_xml)?;

    // Process headers
    let mut header_rel_ids: Vec<(u32, String)> = Vec::new();
    for (header_num, _, _) in &build_result.headers {
        content_types.add_header(*header_num);
        let rel_id = rel_manager.next_id();
        doc_rels.add_header_with_id(&rel_id, *header_num);
        header_rel_ids.push((*header_num, rel_id));
    }

    // Process footers
    let mut footer_rel_ids: Vec<(u32, String)> = Vec::new();
    for (footer_num, _, _) in &build_result.footers {
        content_types.add_footer(*footer_num);
        let rel_id = rel_manager.next_id();
        doc_rels.add_footer_with_id(&rel_id, *footer_num);
        footer_rel_ids.push((*footer_num, rel_id));
    }

    // Update header/footer refs
    for (num, rel_id) in &header_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_header_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_header_id = Some(rel_id.clone());
        } else if *num == 3 {
            // Header 3 is the truly empty header for cover/TOC suppression
            build_result.document.empty_header_id = Some(rel_id.clone());
        }
    }

    for (num, rel_id) in &footer_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_footer_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_footer_id = Some(rel_id.clone());
        } else if *num == 3 {
            // Footer 3 is the truly empty footer for cover/TOC suppression
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

    // Track media files already added to avoid duplicates
    let mut added_media_files: std::collections::HashSet<String> = std::collections::HashSet::new();

    // Add headers
    for (header_num, header_bytes, media) in &build_result.headers {
        packager.add_header(*header_num, header_bytes)?;

        // Add media files for this header (with header_ prefix)
        for (_r_id, media_file) in media {
            if !media_file.filename.is_empty() {
                // Add header_ prefix to avoid conflicts with cover images
                let prefixed_filename = format!("header_{}", media_file.filename);

                // Skip if already added
                if added_media_files.contains(&prefixed_filename) {
                    continue;
                }

                // Add the media file to the archive with prefixed name
                packager.add_image(&prefixed_filename, &media_file.data)?;
                added_media_files.insert(prefixed_filename.clone());

                // Add to content types
                let ext = std::path::Path::new(&prefixed_filename)
                    .extension()
                    .and_then(|s| s.to_str())
                    .unwrap_or("png");
                content_types.add_image_extension(ext, &media_file.content_type);
            }
        }

        // Generate and add rels file if there are media files
        if !media.is_empty() {
            let rels_xml =
                crate::template::render::header_footer::generate_header_footer_rels_xml_with_prefix(
                    media, "header_",
                );
            packager.add_header_rels(*header_num, &rels_xml)?;
        }
    }

    // Add footers
    for (footer_num, footer_bytes, media) in &build_result.footers {
        packager.add_footer(*footer_num, footer_bytes)?;

        // Add media files for this footer (with header_ prefix)
        for (_r_id, media_file) in media {
            if !media_file.filename.is_empty() {
                // Add header_ prefix to avoid conflicts with cover images
                let prefixed_filename = format!("header_{}", media_file.filename);

                // Skip if already added
                if added_media_files.contains(&prefixed_filename) {
                    continue;
                }

                // Add the media file to the archive with prefixed name
                packager.add_image(&prefixed_filename, &media_file.data)?;
                added_media_files.insert(prefixed_filename.clone());

                // Add to content types
                let ext = std::path::Path::new(&prefixed_filename)
                    .extension()
                    .and_then(|s| s.to_str())
                    .unwrap_or("png");
                content_types.add_image_extension(ext, &media_file.content_type);
            }
        }

        // Generate and add rels file if there are media files
        if !media.is_empty() {
            let rels_xml =
                crate::template::render::header_footer::generate_header_footer_rels_xml_with_prefix(
                    media, "header_",
                );
            packager.add_footer_rels(*footer_num, &rels_xml)?;
        }
    }

    let cursor = packager.finish()?;
    Ok(cursor.into_inner())
}

/// Apply cover template to the document
///
/// This function clones the raw XML from cover.docx, replaces placeholders,
/// and inserts it directly into the document. This preserves all original
/// formatting, positions, images, and relationships exactly as designed in Word.
fn apply_cover_template(
    build_result: &mut crate::docx::BuildResult,
    cover: &crate::template::extract::CoverTemplate,
    placeholder_ctx: &crate::template::PlaceholderContext,
    lang: Language,
    rel_manager: &mut crate::docx::RelIdManager,
    table_template: Option<&crate::template::extract::TableTemplate>,
    image_template: Option<&crate::template::extract::ImageTemplate>,
    doc_config: &DocumentConfig,
) -> Result<()> {
    use crate::template::placeholder::replace_placeholders;

    // If we have raw XML from the cover template, use it directly
    if let Some(raw_xml) = &cover.raw_xml {
        // Clone the raw XML
        let mut processed_xml = raw_xml.clone();

        // Handle {{inside}} placeholder specially - it needs markdown rendering
        // Render inside content to XML string
        let inside_xml = if let Some(inside_md) = placeholder_ctx.get("inside") {
            // Parse the inside content as markdown
            let inside_parsed = parse_markdown_with_frontmatter(inside_md);

            // Build the inside content WITHOUT TOC, but WITH page config
            let inside_config = DocumentConfig {
                toc: crate::docx::toc::TocConfig {
                    enabled: false,
                    ..Default::default()
                },
                page: doc_config.page.clone(), // Pass page config to inside content
                ..Default::default()
            };
            // Use the same rel_manager!
            let inside_result = build_document(
                &inside_parsed,
                lang,
                &inside_config,
                rel_manager,
                table_template,
                image_template,
            );

            // Merge resources from inside_result into main build_result
            build_result
                .images
                .images
                .extend(inside_result.images.images);
            build_result
                .hyperlinks
                .hyperlinks
                .extend(inside_result.hyperlinks.hyperlinks);

            // Generate XML string for the inside content
            // We use a temporary DocumentXml to serialize just these elements
            let mut temp_doc = crate::docx::ooxml::DocumentXml::new();
            for elem in inside_result.document.elements {
                // Skip section breaks in inside content to avoid extra pages
                if let crate::docx::ooxml::DocElement::Paragraph(p) = &elem {
                    if p.is_section_break() {
                        continue;
                    }
                }
                temp_doc.elements.push(elem);
            }
            // Use to_xml to get the XML string, but we need to extract body content
            let full_xml = String::from_utf8(temp_doc.to_xml()?).unwrap_or_default();

            // Extract content between <w:body> and </w:body> (excluding sectPr)
            if let Some(start) = full_xml.find("<w:body>") {
                if let Some(end) = full_xml.rfind("<w:sectPr") {
                    // Find last sectPr
                    full_xml[start + 8..end].to_string()
                } else if let Some(end) = full_xml.find("</w:body>") {
                    full_xml[start + 8..end].to_string()
                } else {
                    String::new()
                }
            } else {
                String::new()
            }
        } else {
            String::new()
        };

        // Replace {{inside}} with the rendered XML - DO THIS BEFORE replace_placeholders
        if processed_xml.contains("{{inside}}") {
            if !inside_xml.is_empty() {
                // Try to find the paragraph containing {{inside}} and replace the whole paragraph
                // to avoid nesting paragraphs inside paragraphs (invalid OOXML)
                if let Some(placeholder_pos) = processed_xml.find("{{inside}}") {
                    // Find start of paragraph: look backwards for <w:p> or <w:p ...>
                    // We must avoid matching <w:pPr> or other tags starting with <w:p
                    let slice = &processed_xml[..placeholder_pos];
                    let p_start_1 = slice.rfind("<w:p>");
                    let p_start_2 = slice.rfind("<w:p ");

                    let p_start = match (p_start_1, p_start_2) {
                        (Some(a), Some(b)) => Some(std::cmp::max(a, b)),
                        (Some(a), None) => Some(a),
                        (None, Some(b)) => Some(b),
                        (None, None) => None,
                    };

                    // Find end of paragraph: look forwards for </w:p>
                    let p_end = processed_xml[placeholder_pos..].find("</w:p>");

                    if let (Some(start), Some(end_offset)) = (p_start, p_end) {
                        let end = placeholder_pos + end_offset + 6; // +6 length of </w:p>

                        // Replace the entire paragraph with inside_xml
                        let mut new_xml = String::new();
                        new_xml.push_str(&processed_xml[..start]);
                        new_xml.push_str(&inside_xml);
                        new_xml.push_str(&processed_xml[end..]);
                        processed_xml = new_xml;
                    } else {
                        // Fallback: simple text replacement
                        processed_xml = processed_xml.replace("{{inside}}", &inside_xml);
                    }
                }
            } else {
                // Remove {{inside}} if no content provided
                processed_xml = processed_xml.replace("{{inside}}", "");
            }
        }

        // Replace other simple placeholders (like {{title}}, {{author}})
        processed_xml = replace_placeholders(&processed_xml, placeholder_ctx);

        // Fix image relationship IDs
        // Map old rId (from cover.docx) to new rId (in generated docx)
        let mut processed_filenames: std::collections::HashSet<String> =
            std::collections::HashSet::new();

        // Process all images from cover.elements (now includes SVG companion files)
        for element in &cover.elements {
            if let crate::template::extract::CoverElement::Image {
                rel_id,
                filename,
                data,
                width,
                height,
                ..
            } = element
            {
                if let Some(img_data) = data {
                    // Generate new relationship ID using RelIdManager
                    let new_rel_id = rel_manager.get_mapped_id("cover", rel_id);

                    // Replace old ID with new ID in the XML
                    processed_xml = processed_xml.replace(
                        &format!("r:embed=\"{}\"", rel_id),
                        &format!("r:embed=\"{}\"", new_rel_id),
                    );

                    // Check if image with this filename already exists to avoid duplicates
                    let already_exists = processed_filenames.contains(filename)
                        || build_result
                            .images
                            .images
                            .iter()
                            .any(|img| img.filename == *filename);
                    if !already_exists {
                        processed_filenames.insert(filename.clone());
                        // Add image to build result
                        let img_info = crate::docx::ImageInfo {
                            filename: filename.clone(),
                            rel_id: new_rel_id.clone(),
                            src: filename.clone(),
                            data: Some(img_data.clone()),
                            width_emu: *width,
                            height_emu: *height,
                        };
                        build_result.images.images.push(img_info);
                    }
                }
            }
        }

        // Strip any <w:sectPr> from the cover XML - we want to control page layout ourselves
        // The sectPr in the cover template would override our section break settings
        processed_xml = strip_section_properties(&processed_xml);

        // Add the processed raw XML as a DocElement
        // Insert at the beginning of the document
        build_result
            .document
            .elements
            .insert(0, crate::docx::ooxml::DocElement::RawXml(processed_xml));

        // Add a section break after the cover to separate it from TOC/content
        // Apply page config if available
        let mut cover_section_break = crate::docx::ooxml::Paragraph::new()
            .section_break("nextPage")
            .suppress_header_footer();

        // Apply page layout from config
        if let Some(ref page_config) = doc_config.page {
            cover_section_break = cover_section_break.with_page_layout(
                page_config.width,
                page_config.height,
                page_config.margin_top,
                page_config.margin_right,
                page_config.margin_bottom,
                page_config.margin_left,
                page_config.margin_header,
                page_config.margin_footer,
                page_config.margin_gutter,
            );
        }

        build_result.document.elements.insert(
            1,
            crate::docx::ooxml::DocElement::Paragraph(Box::new(cover_section_break)),
        );
    }

    Ok(())
}

/// Strip <w:sectPr> elements from XML string
///
/// This removes section properties from the cover template XML so that
/// we can control page layout through our own section break.
fn strip_section_properties(xml: &str) -> String {
    let mut result = xml.to_string();

    // Remove <w:sectPr>...</w:sectPr> elements
    // Handle both self-closing and content forms
    loop {
        // Find the start of a sectPr element
        if let Some(start) = result.find("<w:sectPr") {
            // Find the end of this element
            // It could be self-closing: <w:sectPr ... /> or have content: <w:sectPr ...>...</w:sectPr>
            let after_start = &result[start..];

            // Check if it's self-closing
            if let Some(self_close) = after_start.find("/>") {
                let open_end = after_start.find('>').unwrap_or(self_close);
                if self_close == open_end - 1 {
                    // It's self-closing: <w:sectPr ... />
                    result.replace_range(start..start + self_close + 2, "");
                    continue;
                }
            }

            // It's a container element, find the closing tag
            if let Some(end) = result[start..].find("</w:sectPr>") {
                result.replace_range(start..start + end + 11, "");
                continue;
            }

            // If we can't find the end, break to avoid infinite loop
            break;
        } else {
            break;
        }
    }

    result
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

    let mut build_result = build_document(
        &parsed,
        lang,
        &DocumentConfig::default(),
        &mut crate::docx::RelIdManager::new(),
        None,
        None,
    );

    // Insert TOC if enabled
    if let Some(toc_builder) = build_result.toc_builder.take() {
        let toc_config = TocConfig::default();
        if toc_config.enabled && !toc_builder.is_empty() {
            let toc_elements = toc_builder.generate_toc(&toc_config);
            // Prepend TOC at the beginning
            for (i, elem) in toc_elements.into_iter().enumerate() {
                build_result.document.elements.insert(i, elem);
            }
        }
    }

    let buffer = Cursor::new(Vec::new());
    let mut packager = Packager::new(buffer);

    let mut content_types = ContentTypes::new();
    let rels = Relationships::root_rels();
    let mut doc_rels = Relationships::document_rels();
    let styles = StylesDocument::new(lang, None);

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
            if let Some(ref data) = image.data {
                packager.add_image(&image.filename, data)?;
            } else if let Ok(data) = std::fs::read(&image.src) {
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
    for (header_num, _, _) in &build_result.headers {
        content_types.add_header(*header_num);
        let rel_id = doc_rels.add_header(*header_num);
        header_rel_ids.push((*header_num, rel_id));
    }

    // Process footers and capture returned relationship IDs
    let mut footer_rel_ids: Vec<(u32, String)> = Vec::new();
    for (footer_num, _, _) in &build_result.footers {
        content_types.add_footer(*footer_num);
        let rel_id = doc_rels.add_footer(*footer_num);
        footer_rel_ids.push((*footer_num, rel_id));
    }

    // Update header/footer refs with actual relationship IDs from doc_rels
    // Header 1 is default, header 2 is first page, header 3 is empty for suppression
    for (num, rel_id) in &header_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_header_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_header_id = Some(rel_id.clone());
        } else if *num == 3 {
            // Header 3 is the truly empty header for cover/TOC suppression
            build_result.document.empty_header_id = Some(rel_id.clone());
        }
    }

    for (num, rel_id) in &footer_rel_ids {
        if *num == 1 {
            build_result.document.header_footer_refs.default_footer_id = Some(rel_id.clone());
        } else if *num == 2 {
            build_result.document.header_footer_refs.first_footer_id = Some(rel_id.clone());
        } else if *num == 3 {
            // Footer 3 is the truly empty footer for cover/TOC suppression
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

    // Track media files already added to avoid duplicates
    let mut added_media_files: std::collections::HashSet<String> = std::collections::HashSet::new();

    // Add headers to the archive
    for (header_num, header_bytes, media) in &build_result.headers {
        packager.add_header(*header_num, header_bytes)?;

        // Add media files for this header (with header_ prefix)
        for (_r_id, media_file) in media {
            if !media_file.filename.is_empty() {
                // Add header_ prefix to avoid conflicts with cover images
                let prefixed_filename = format!("header_{}", media_file.filename);

                // Skip if already added
                if added_media_files.contains(&prefixed_filename) {
                    continue;
                }

                // Add the media file to the archive with prefixed name
                packager.add_image(&prefixed_filename, &media_file.data)?;
                added_media_files.insert(prefixed_filename.clone());

                // Add to content types
                let ext = std::path::Path::new(&prefixed_filename)
                    .extension()
                    .and_then(|s| s.to_str())
                    .unwrap_or("png");
                content_types.add_image_extension(ext, &media_file.content_type);
            }
        }

        // Generate and add rels file if there are media files
        if !media.is_empty() {
            let rels_xml =
                crate::template::render::header_footer::generate_header_footer_rels_xml_with_prefix(
                    media, "header_",
                );
            packager.add_header_rels(*header_num, &rels_xml)?;
        }
    }

    // Add footers to the archive
    for (footer_num, footer_bytes, media) in &build_result.footers {
        packager.add_footer(*footer_num, footer_bytes)?;

        // Add media files for this footer (with header_ prefix)
        for (_r_id, media_file) in media {
            if !media_file.filename.is_empty() {
                // Add header_ prefix to avoid conflicts with cover images
                let prefixed_filename = format!("header_{}", media_file.filename);

                // Skip if already added
                if added_media_files.contains(&prefixed_filename) {
                    continue;
                }

                // Add the media file to the archive with prefixed name
                packager.add_image(&prefixed_filename, &media_file.data)?;
                added_media_files.insert(prefixed_filename.clone());

                // Add to content types
                let ext = std::path::Path::new(&prefixed_filename)
                    .extension()
                    .and_then(|s| s.to_str())
                    .unwrap_or("png");
                content_types.add_image_extension(ext, &media_file.content_type);
            }
        }

        // Generate and add rels file if there are media files
        if !media.is_empty() {
            let rels_xml =
                crate::template::render::header_footer::generate_header_footer_rels_xml_with_prefix(
                    media, "header_",
                );
            packager.add_footer_rels(*footer_num, &rels_xml)?;
        }
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
        let mut rel_manager = crate::docx::RelIdManager::new();
        let build_result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        );

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
        let styles = StylesDocument::new(Language::English, None);

        // Add header/footer content types (header1, footer1)
        content_types.add_header(1);
        content_types.add_footer(1);

        // Add header/footer relationships (returns the relationship IDs)
        let header_rel_id = rel_manager.next_id();
        doc_rels.add_header_with_id(&header_rel_id, 1);

        let footer_rel_id = rel_manager.next_id();
        doc_rels.add_footer_with_id(&footer_rel_id, 1);

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
        for (header_num, header_bytes, _) in &build_result.headers {
            packager.add_header(*header_num, header_bytes).unwrap();
        }

        for (footer_num, footer_bytes, _) in &build_result.footers {
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
        let mut rel_manager = crate::docx::RelIdManager::new();
        let build_result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        );

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
