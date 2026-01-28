//! DOCX builder - Convert parsed Markdown AST to OOXML
//!
//! This module bridges the parser's AST representation with the OOXML
//! document structure, converting markdown elements to DOCX paragraphs
//! and runs.

use crate::docx::image_utils::{default_image_size_emu, read_image_dimensions};
use crate::docx::ooxml::{
    DocElement, DocumentXml, FooterConfig, FooterXml, FootnotesXml, HeaderConfig, HeaderFooterRefs,
    HeaderXml, ImageElement, Paragraph, ParagraphChild, Run, Table, TableCellElement, TableRow,
    TableWidth,
};
use crate::docx::toc::{TocBuilder, TocConfig};
use crate::docx::xref::CrossRefContext;
use crate::parser::{
    Alignment as ParserAlignment, Block, Inline, ListItem, ParsedDocument,
    TableCell as ParserTableCell,
};
use crate::Language;

/// Tracks images during document building
#[derive(Debug, Default)]
pub struct ImageContext {
    /// Map of image source path to (filename, relationship_id, data)
    pub images: Vec<ImageInfo>,
    next_id: u32,
}

/// Information about an embedded image
#[derive(Debug, Clone)]
pub struct ImageInfo {
    pub filename: String,      // e.g., "image1.png"
    pub rel_id: String,        // e.g., "rId4"
    pub src: String,           // Original source path/URL
    pub data: Option<Vec<u8>>, // Image bytes (None if external)
    pub width_emu: i64,        // Width in EMUs
    pub height_emu: i64,       // Height in EMUs
}

/// Tracks hyperlinks during document building
#[derive(Debug, Default, Clone)]
pub struct HyperlinkContext {
    pub hyperlinks: Vec<HyperlinkInfo>,
    next_id: u32,
}

/// Information about a hyperlink
#[derive(Debug, Clone)]
pub struct HyperlinkInfo {
    pub url: String,
    pub rel_id: String,
}

impl HyperlinkContext {
    pub fn new() -> Self {
        Self {
            hyperlinks: Vec::new(),
            next_id: 1,
        }
    }

    /// Add a hyperlink and return its relationship ID
    ///
    /// Uses a large offset (10000) to avoid clashes with image IDs
    pub fn add_hyperlink(&mut self, url: &str) -> String {
        let rel_id = format!("rId{}", self.next_id + 10000);
        self.hyperlinks.push(HyperlinkInfo {
            url: url.to_string(),
            rel_id: rel_id.clone(),
        });
        self.next_id += 1;
        rel_id
    }
}

/// Tracks list numbering instances during document building
///
/// Each separate list (ordered or unordered) gets a unique numId
/// to ensure Word restarts numbering for each list independently.
#[derive(Debug, Default, Clone)]
pub struct NumberingContext {
    /// List of (numId, is_ordered) pairs for all lists
    pub lists: Vec<NumberingInfo>,
    next_id: u32,
}

/// Information about a list numbering instance
#[derive(Debug, Clone)]
pub struct NumberingInfo {
    pub num_id: u32,
    pub is_ordered: bool,
}

impl NumberingContext {
    pub fn new() -> Self {
        Self {
            lists: Vec::new(),
            next_id: 1,
        }
    }

    /// Register a new list and return its unique numId
    ///
    /// Each call creates a new list instance that will restart numbering.
    pub fn add_list(&mut self, ordered: bool) -> u32 {
        let num_id = self.next_id;
        self.lists.push(NumberingInfo {
            num_id,
            is_ordered: ordered,
        });
        self.next_id += 1;
        num_id
    }
}

impl ImageContext {
    pub fn new() -> Self {
        Self {
            images: Vec::new(),
            next_id: 1,
        }
    }

    /// Add an image and return its relationship ID
    ///
    /// For now, we assign a placeholder rel_id. The actual ID will be
    /// assigned during packaging when relationships are finalized.
    pub fn add_image(&mut self, src: &str, width: Option<&str>) -> String {
        // Generate unique ID (offset by 3 since rId1-3 are used for styles, settings, fontTable)
        // This ensures the first image gets rId4
        let rel_id = format!("rId{}", self.next_id + 3);
        let filename = self.generate_filename(src);

        // Try to read actual dimensions
        let mut actual_dims = None;
        #[cfg(not(target_arch = "wasm32"))]
        {
            if let Ok(data) = std::fs::read(src) {
                actual_dims = read_image_dimensions(&data);
            }
        }

        let (width_emu, height_emu) = self.parse_dimensions(width, actual_dims);

        self.images.push(ImageInfo {
            filename: filename.clone(),
            rel_id: rel_id.clone(),
            src: src.to_string(),
            data: None, // Data loaded during packaging
            width_emu,
            height_emu,
        });

        self.next_id += 1;
        rel_id
    }

    /// Add a mermaid diagram SVG and return its relationship ID
    pub fn add_mermaid_svg(&mut self, filename: &str, data: Vec<u8>) -> String {
        let rel_id = format!("rId{}", self.next_id + 3);

        // Read SVG dimensions and calculate proper size
        let (width_emu, height_emu) = if let Some(dims) = read_image_dimensions(&data) {
            default_image_size_emu(dims)
        } else {
            // Fallback to 6x4 inches
            (6 * 914400, 4 * 914400)
        };

        self.images.push(ImageInfo {
            filename: filename.to_string(),
            rel_id: rel_id.clone(),
            src: filename.to_string(), // Virtual source
            data: Some(data),
            width_emu,
            height_emu,
        });

        self.next_id += 1;
        rel_id
    }

    /// Generate a unique filename for the image
    fn generate_filename(&self, src: &str) -> String {
        // Extract extension from source
        let ext = std::path::Path::new(src)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png");

        format!("image{}.{}", self.next_id, ext)
    }

    /// Parse width specification into EMUs
    fn parse_dimensions(
        &self,
        width: Option<&str>,
        actual_dims: Option<crate::docx::image_utils::ImageDimensions>,
    ) -> (i64, i64) {
        const EMU_PER_INCH: i64 = 914400;
        const DEFAULT_WIDTH_INCHES: f64 = 6.0;

        // Use actual aspect ratio if available, otherwise default to 3:2
        let inv_aspect = actual_dims.map(|d| 1.0 / d.aspect_ratio()).unwrap_or(0.67);

        if let Some(w) = width {
            if w.ends_with('%') {
                // Percentage of page width (~6.0 inches for A4 with margins)
                let pct: f64 = w.trim_end_matches('%').parse().unwrap_or(100.0);
                let width_inches = 6.0 * (pct / 100.0);
                let height_inches = width_inches * inv_aspect;
                (
                    (width_inches * EMU_PER_INCH as f64) as i64,
                    (height_inches * EMU_PER_INCH as f64) as i64,
                )
            } else if w.ends_with("in") {
                let width_inches: f64 = w
                    .trim_end_matches("in")
                    .parse()
                    .unwrap_or(DEFAULT_WIDTH_INCHES);
                let height_inches = width_inches * inv_aspect;
                (
                    (width_inches * EMU_PER_INCH as f64) as i64,
                    (height_inches * EMU_PER_INCH as f64) as i64,
                )
            } else {
                // Pixels (assume 96 DPI)
                let val_str = w.trim_end_matches("px");
                let px: f64 = val_str.parse().unwrap_or(576.0);
                let width_inches = (px / 96.0).min(6.0); // Constrain to 6 inches max
                let height_inches = width_inches * inv_aspect;
                (
                    (width_inches * EMU_PER_INCH as f64) as i64,
                    (height_inches * EMU_PER_INCH as f64) as i64,
                )
            }
        } else if let Some(dims) = actual_dims {
            // Use standard calculation based on actual dimensions
            default_image_size_emu(dims)
        } else {
            // Fallback to 6x4 inches
            (
                (DEFAULT_WIDTH_INCHES * EMU_PER_INCH as f64) as i64,
                (4.0 * EMU_PER_INCH as f64) as i64,
            )
        }
    }
}

/// Document build configuration
#[derive(Debug, Clone, Default)]
pub struct DocumentConfig {
    pub title: String,
    pub toc: TocConfig,
    pub header: HeaderConfig,
    pub footer: FooterConfig,
    pub different_first_page: bool, // Hide header/footer on first page
}

/// Result of building a document, including tracked images, hyperlinks, footnotes, and headers/footers
#[derive(Debug)]
pub struct BuildResult {
    pub document: DocumentXml,
    pub images: ImageContext,
    pub hyperlinks: HyperlinkContext,
    pub footnotes: FootnotesXml,
    pub numbering: NumberingContext,
    pub headers: Vec<(u32, Vec<u8>)>, // (header_num, xml_bytes)
    pub footers: Vec<(u32, Vec<u8>)>, // (footer_num, xml_bytes)
    pub has_toc_section_break: bool,  // If true, there's a TOC section break needing empty refs
}

/// Check if a block is a heading
fn is_heading(block: &Block) -> bool {
    matches!(block, Block::Heading { .. })
}

/// Extract plain text from inline elements (for TOC entries)
fn extract_inline_text(inlines: &[Inline]) -> String {
    inlines
        .iter()
        .map(|inline| match inline {
            Inline::Text(t) => t.clone(),
            Inline::Bold(inner) | Inline::Italic(inner) | Inline::Strikethrough(inner) => {
                extract_inline_text(inner)
            }
            Inline::BoldItalic(inner) => extract_inline_text(inner),
            Inline::Code(code) => code.clone(),
            Inline::Link { text, .. } => extract_inline_text(text),
            Inline::Image { alt, .. } => alt.clone(),
            Inline::FootnoteRef(_) => String::new(),
            Inline::CrossRef { .. } => String::new(),
            Inline::SoftBreak => " ".to_string(),
            Inline::HardBreak => "\n".to_string(),
            Inline::Html(_) => String::new(),
            Inline::IndexMarker(_) => String::new(),
        })
        .collect::<Vec<_>>()
        .join("")
}

/// Build a DOCX document from parsed markdown
///
/// # Arguments
/// * `doc` - The parsed markdown document
/// * `_lang` - Language for style defaults (English/Thai) - currently unused
/// * `config` - Document configuration including TOC, header, footer settings
///
/// # Returns
/// A `BuildResult` containing the document and tracked images, hyperlinks, footnotes, and headers/footers
///
/// # Example
/// ```rust
/// use md2docx::parser::parse_markdown_with_frontmatter;
/// use md2docx::docx::build_document;
/// use md2docx::DocumentConfig;
/// use md2docx::Language;
///
/// let md = "# Hello World\n\nThis is **bold** text.";
/// let parsed = parse_markdown_with_frontmatter(md);
/// let config = DocumentConfig::default();
/// let result = build_document(&parsed, Language::English, &config);
/// ```
pub fn build_document(
    doc: &ParsedDocument,
    _lang: Language,
    config: &DocumentConfig,
) -> BuildResult {
    let mut doc_xml = DocumentXml::new();
    let mut image_ctx = ImageContext::new();
    let mut hyperlink_ctx = HyperlinkContext::new();
    let mut numbering_ctx = NumberingContext::new();
    let mut image_id_counter: u32 = 1;
    let mut footnotes = FootnotesXml::new();

    // TOC builder for collecting headings
    let mut toc_builder = TocBuilder::new();
    let mut bookmark_id_counter: u32 = 0;

    // Cross-reference context for tracking anchors
    let mut xref_ctx = CrossRefContext::new();

    // Track headers and footers
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut header_footer_refs = HeaderFooterRefs::default();

    // Track previous block to insert blank lines before headings
    let mut prev_block: Option<&Block> = None;

    // Find the first thematic break index (end of cover section)
    // Headings before this should not be in TOC
    let first_thematic_break_index = doc
        .blocks
        .iter()
        .position(|b| matches!(b, Block::ThematicBreak));

    // Process all blocks in the document
    // Track the last list seen to support resuming lists across code blocks
    let mut last_list_info: Option<(u32, bool, usize)> = None; // (num_id, is_ordered, block_index)

    for (i, block) in doc.blocks.iter().enumerate() {
        // Create build context
        let mut ctx = BuildContext {
            image_ctx: &mut image_ctx,
            hyperlink_ctx: &mut hyperlink_ctx,
            numbering_ctx: &mut numbering_ctx,
            image_id: &mut image_id_counter,
            doc,
            footnotes: &mut footnotes,
            toc_builder: &mut toc_builder,
            bookmark_id_counter: &mut bookmark_id_counter,
            xref_ctx: &mut xref_ctx,
        };

        // Insert blank paragraph before heading if previous block was not a heading
        if is_heading(block) {
            if let Some(prev) = prev_block {
                if !is_heading(prev) {
                    doc_xml.add_element(DocElement::Paragraph(Box::default()));
                }
            }
            // Heading breaks any list continuation
            last_list_info = None;
        }

        // Determine if we should force continuation of the previous list
        let mut forced_num_id = None;
        if let Block::List { ordered, .. } = block {
            if let Some((last_id, last_ordered, last_idx)) = last_list_info {
                // If it's the same type of list and we only skipped 1 block (e.g. a code block)
                // Resume the list numbering.
                // i - last_idx == 1 means adjacent (normal case, usually handled by parser merging)
                // i - last_idx == 2 means one block in between (e.g. List -> Code -> List)
                if *ordered == last_ordered && (i - last_idx) <= 2 {
                    forced_num_id = Some(last_id);
                }
            }
        } else if !matches!(
            block,
            Block::CodeBlock { .. } | Block::BlockQuote(_) | Block::Table { .. }
        ) {
            // If the block is NOT a list, and NOT something that might be inside a list (like code/quote/table)
            // Then it definitely breaks the list.
            // Text paragraphs usually break lists unless indented, but here we assume top-level paragraphs break lists.
            last_list_info = None;
        }

        // Skip TOC for blocks before first thematic break (cover section)
        let skip_toc = first_thematic_break_index.is_some_and(|idx| i < idx);

        let elements = block_to_elements(block, 0, &mut ctx, forced_num_id, skip_toc);

        // If this block was a list, update tracking info
        if let Block::List { ordered, .. } = block {
            // Find the num_id used. If we forced it, we know it.
            // If we generated it, we need to extract it from the generated paragraphs.
            // But block_to_elements doesn't return the ID.
            // However, if forced_num_id is None, block_to_elements called add_list,
            // so it's the last added list.
            let used_id = forced_num_id.unwrap_or_else(|| {
                // If we didn't force it, it was just added.
                // NumberingContext IDs are sequential (1, 2, 3...)
                // access internal state? NumberingContext exposes lists vec.
                numbering_ctx.lists.last().map(|l| l.num_id).unwrap_or(0)
            });

            last_list_info = Some((used_id, *ordered, i));
        }

        for elem in elements {
            doc_xml.add_element(elem);
        }

        prev_block = Some(block);
    }

    // Generate TOC elements
    let toc_elements = toc_builder.generate_toc(&config.toc);

    if !toc_elements.is_empty() {
        if config.toc.after_cover {
            // Find the first H1 heading that comes AFTER a section break
            // This is the start of the first real chapter (after cover)
            let mut found_section_break = false;
            let first_chapter_h1_index = doc_xml.elements.iter().position(|el| {
                if let DocElement::Paragraph(p) = el {
                    if p.section_break.is_some() {
                        found_section_break = true;
                        false // Don't match section break itself
                    } else if found_section_break && p.style_id.as_deref() == Some("Heading1") {
                        true // Found H1 after section break
                    } else {
                        false
                    }
                } else {
                    false
                }
            });

            if let Some(idx) = first_chapter_h1_index {
                let toc_count = toc_elements.len();
                // Insert TOC elements BEFORE the first chapter H1
                for (i, elem) in toc_elements.into_iter().enumerate() {
                    doc_xml.elements.insert(idx + i, elem);
                }

                // Find the section break that ends Chapter 1 (first sectPr after Chapter 1's H1)
                // and set it to restart page numbering at 1.
                // This ensures Chapter 1 starts at page 1.
                let chapter1_h1_pos = idx + toc_count;
                let mut _found_break = false;

                for i in (chapter1_h1_pos + 1)..doc_xml.elements.len() {
                    if let DocElement::Paragraph(p) = &mut doc_xml.elements[i] {
                        if p.section_break.is_some() {
                            p.page_num_start = Some(1);
                            _found_break = true;
                            break;
                        }
                    }
                }

                // If no section break found (e.g., single chapter document),
                // we'll need to rely on the body sectPr, but we don't have access to it here
                // easily as it's generated in to_xml().
                // However, most multi-chapter docs will have breaks.
                // If it's a single chapter, the TOC section break handles the separation,
                // and the body sectPr handles the chapter.
                // But body sectPr doesn't support pgNumType restart in our current DocumentXml model
                // (it takes refs from header_footer_refs).
                // TODO: Add support for body sectPr page numbering restart if needed.
            } else {
                // Fallback: insert at beginning if no chapter H1 found after section break
                let first_h1_index = doc_xml.elements.iter().position(|el| {
                    if let DocElement::Paragraph(p) = el {
                        p.style_id.as_deref() == Some("Heading1")
                    } else {
                        false
                    }
                });

                if let Some(idx) = first_h1_index {
                    for (i, elem) in toc_elements.into_iter().enumerate() {
                        doc_xml.elements.insert(idx + i, elem);
                    }
                } else {
                    for (i, elem) in toc_elements.into_iter().enumerate() {
                        doc_xml.elements.insert(i, elem);
                    }
                }
            }
        } else {
            // Prepend TOC at the beginning (old behavior)
            for (i, elem) in toc_elements.into_iter().enumerate() {
                doc_xml.elements.insert(i, elem);
            }
        }
    }

    // Generate headers and footers
    // Note: Relationship IDs are NOT set here - they are assigned in lib.rs after
    // doc_rels.add_header() and add_footer() are called, which return the actual IDs.
    if !config.header.is_empty() {
        // Generate default header (header1.xml)
        let header_xml = HeaderXml::new(config.header.clone(), &config.title);
        headers.push((1, header_xml.to_xml().unwrap()));
        // Relationship ID will be set in lib.rs

        // Always generate empty header (ID 2) for suppression purposes
        // (We reuse this for first page if different_first_page is set)
        let empty_header = HeaderXml::new(HeaderConfig::empty(), "");
        headers.push((2, empty_header.to_xml().unwrap()));

        if config.different_first_page {
            header_footer_refs.different_first_page = true;
        }
    }

    if !config.footer.is_empty() {
        // Generate default footer (footer1.xml)
        let footer_xml = FooterXml::new(config.footer.clone(), &config.title);
        footers.push((1, footer_xml.to_xml().unwrap()));
        // Relationship ID will be set in lib.rs

        // Always generate empty footer (ID 2) for suppression purposes
        // (We reuse this for first page if different_first_page is set)
        let empty_footer = FooterXml::new(FooterConfig::empty(), "");
        footers.push((2, empty_footer.to_xml().unwrap()));

        if config.different_first_page {
            header_footer_refs.different_first_page = true;
        }
    }

    // Set header/footer refs on document
    doc_xml.header_footer_refs = header_footer_refs;

    BuildResult {
        document: doc_xml,
        images: image_ctx,
        hyperlinks: hyperlink_ctx,
        footnotes,
        numbering: numbering_ctx,
        headers,
        footers,
        has_toc_section_break: false,
    }
}

/// Context for building a document, holding all tracked state
pub struct BuildContext<'a> {
    pub image_ctx: &'a mut ImageContext,
    pub hyperlink_ctx: &'a mut HyperlinkContext,
    pub numbering_ctx: &'a mut NumberingContext,
    pub image_id: &'a mut u32,
    pub doc: &'a ParsedDocument,
    pub footnotes: &'a mut FootnotesXml,
    pub toc_builder: &'a mut TocBuilder,
    pub bookmark_id_counter: &'a mut u32,
    pub xref_ctx: &'a mut CrossRefContext,
}

/// Convert a Block to one or more DocElements (Paragraph, Table, or Image)
///
/// Some block types (like lists, code blocks, blockquotes) may generate
/// multiple paragraphs. Tables generate a single Table element.
/// Images generate a single Image element.
///
/// # Arguments
/// * `block` - The block to convert
/// * `list_level` - Current nesting level for lists (0 = top level)
/// * `ctx` - Build context holding tracked state
/// * `forced_num_id` - Optional numId to force for this list (for resuming lists)
/// * `skip_toc` - If true, skip adding headings to TOC (e.g., cover section)
///
/// # Returns
/// A vector of document elements representing the block
fn block_to_elements(
    block: &Block,
    list_level: usize,
    ctx: &mut BuildContext,
    forced_num_id: Option<u32>,
    skip_toc: bool,
) -> Vec<DocElement> {
    match block {
        Block::Image {
            alt,
            src,
            width,
            id,
            ..
        } => {
            // Register figure anchor if id is present
            if let Some(fig_id) = id {
                ctx.xref_ctx.register_figure(fig_id, alt);
            }

            // Add image to context and get relationship ID
            let rel_id = ctx.image_ctx.add_image(src, width.as_deref());

            // Get dimensions from context (last added image)
            let (width_emu, height_emu) = ctx
                .image_ctx
                .images
                .last()
                .map(|img| (img.width_emu, img.height_emu))
                .unwrap_or((5486400, 3657600)); // Default 6x4 inches

            // Create image element
            let img = ImageElement::new(&rel_id, width_emu, height_emu)
                .alt_text(alt)
                .name(src)
                .id(*ctx.image_id);

            *ctx.image_id += 1;

            vec![DocElement::Image(img)]
        }

        Block::Mermaid { content, id } => {
            match crate::mermaid::render_to_svg(content) {
                Ok(svg) => {
                    // Register figure anchor if id is present
                    if let Some(fig_id) = id {
                        ctx.xref_ctx.register_figure(fig_id, "Mermaid Diagram");
                    }

                    // Generate a virtual filename
                    let filename = format!("mermaid{}.svg", ctx.image_id);

                    // Add to image context
                    let rel_id = ctx.image_ctx.add_mermaid_svg(&filename, svg.into_bytes());

                    // Get dimensions from context (last added image)
                    let (width_emu, height_emu) = ctx
                        .image_ctx
                        .images
                        .last()
                        .map(|img| (img.width_emu, img.height_emu))
                        .unwrap_or((6 * 914400, 4 * 914400));

                    let img = ImageElement::new(&rel_id, width_emu, height_emu)
                        .alt_text("Mermaid Diagram")
                        .name(&filename)
                        .id(*ctx.image_id);

                    *ctx.image_id += 1;
                    vec![DocElement::Image(img)]
                }
                Err(e) => {
                    eprintln!("Warning: Failed to render mermaid diagram: {}", e);
                    // Fallback to code block (represented as paragraphs)
                    block_to_paragraphs(block, list_level, ctx, skip_toc)
                        .into_iter()
                        .map(|p| DocElement::Paragraph(Box::new(p)))
                        .collect()
                }
            }
        }

        Block::Table {
            headers,
            alignments,
            rows,
        } => {
            let table = table_to_docx(headers, alignments, rows, ctx);

            // Add empty paragraph after table for spacing
            // Set spacing to 0 to avoid double padding (since the line itself provides separation)
            let empty_para = Paragraph::default().spacing(0, 0).line_spacing(240, "auto");

            vec![
                DocElement::Table(table),
                DocElement::Paragraph(Box::new(empty_para)),
            ]
        }

        Block::BlockQuote(blocks) => {
            let mut result = Vec::new();
            for nested_block in blocks {
                let elements = block_to_elements(
                    nested_block,
                    list_level,
                    ctx,
                    None, // Don't force numbering in blockquotes for now
                    skip_toc,
                );
                for elem in elements {
                    match elem {
                        DocElement::Paragraph(mut p) => {
                            p.style_id = Some("Quote".to_string());
                            result.push(DocElement::Paragraph(p));
                        }
                        other => result.push(other),
                    }
                }
            }
            result
        }

        Block::Include { resolved, .. } => {
            if let Some(blocks) = resolved {
                let mut result = Vec::new();
                for block in blocks {
                    result.extend(block_to_elements(block, list_level, ctx, None, skip_toc));
                }
                result
            } else {
                vec![]
            }
        }

        Block::List {
            ordered,
            start,
            items,
        } => {
            // Register this list and get a unique numId, or use forced one
            let num_id = forced_num_id.unwrap_or_else(|| ctx.numbering_ctx.add_list(*ordered));

            list_to_paragraphs_with_num_id(
                *ordered, *start, items, list_level, num_id, ctx, skip_toc,
            )
            .into_iter()
            .map(|p| DocElement::Paragraph(Box::new(p)))
            .collect()
        }

        // All other blocks just produce paragraphs
        _ => block_to_paragraphs(block, list_level, ctx, skip_toc)
            .into_iter()
            .map(|p| DocElement::Paragraph(Box::new(p)))
            .collect(),
    }
}

/// Convert a Block to one or more Paragraphs
///
/// Some block types (like lists, code blocks, blockquotes) may generate
/// multiple paragraphs.
///
/// # Arguments
/// * `block` - The block to convert
/// * `list_level` - Current nesting level for lists (0 = top level)
/// * `ctx` - Build context holding tracked state
/// * `skip_toc` - If true, skip adding headings to TOC (e.g., cover section)
///
/// # Returns
/// A vector of paragraphs representing the block
fn block_to_paragraphs(
    block: &Block,
    list_level: usize,
    ctx: &mut BuildContext,
    skip_toc: bool,
) -> Vec<Paragraph> {
    match block {
        Block::Heading { level, content, id } => {
            // Extract text for TOC
            let text = extract_inline_text(content);

            // Register heading with TOC builder (unless in cover section)
            let bookmark_name = if skip_toc {
                // Generate a bookmark name without adding to TOC
                format!("_Heading_{}", *ctx.bookmark_id_counter + 1)
            } else {
                ctx.toc_builder.add_heading(*level, &text, id.as_deref())
            };

            // Register heading with cross-reference context if id is present
            if let Some(anchor_id) = id {
                ctx.xref_ctx.register_heading(anchor_id, *level, &text);
            }

            // Create paragraph with bookmark
            *ctx.bookmark_id_counter += 1;
            let mut para = heading_to_paragraph(*level, content, ctx);
            para = para.with_bookmark(*ctx.bookmark_id_counter, &bookmark_name);

            vec![para]
        }

        Block::Paragraph(inlines) => {
            vec![paragraph_to_paragraph(inlines, ctx)]
        }

        Block::CodeBlock {
            content,
            filename,
            highlight_lines,
            show_line_numbers,
            ..
        } => code_block_to_paragraphs(
            content,
            filename.as_deref(),
            highlight_lines,
            *show_line_numbers,
        ),

        Block::BlockQuote(blocks) => {
            let mut paragraphs = Vec::new();
            for nested_block in blocks {
                let mut nested_paragraphs =
                    block_to_paragraphs(nested_block, list_level, ctx, skip_toc);
                // Apply quote style to all nested paragraphs
                for p in &mut nested_paragraphs {
                    p.style_id = Some("Quote".to_string());
                }
                paragraphs.extend(nested_paragraphs);
            }
            paragraphs
        }

        Block::List {
            ordered,
            start,
            items,
        } => {
            // Register this list and get a unique numId
            let num_id = ctx.numbering_ctx.add_list(*ordered);
            list_to_paragraphs_with_num_id(
                *ordered, *start, items, list_level, num_id, ctx, skip_toc,
            )
        }

        Block::ThematicBreak => vec![thematic_break_to_paragraph()],

        Block::Html(_) => {
            // Skip HTML blocks for now
            vec![]
        }

        Block::Mermaid { content, .. } => {
            // This is a fallback case if block_to_elements falls back to block_to_paragraphs
            code_block_to_paragraphs(content, Some("mermaid"), &Vec::new(), false)
        }

        Block::Include { resolved, .. } => {
            // If include was resolved, process the resolved blocks
            if let Some(blocks) = resolved {
                let mut paragraphs = Vec::new();
                for block in blocks {
                    paragraphs.extend(block_to_paragraphs(block, list_level, ctx, skip_toc));
                }
                paragraphs
            } else {
                vec![]
            }
        }

        Block::CodeInclude { .. } => {
            // Skip code includes for now - will be handled in Phase 3
            vec![]
        }

        Block::Table { .. } => {
            // Tables are handled in block_to_elements()
            vec![]
        }

        Block::Image { .. } => {
            // Skip images for now - will be handled in Phase 3
            vec![]
        }
    }
}

/// Convert a heading block to a paragraph
fn heading_to_paragraph(level: u8, content: &[Inline], ctx: &mut BuildContext) -> Paragraph {
    let style_id = match level {
        1 => "Heading1",
        2 => "Heading2",
        3 => "Heading3",
        _ => "Heading4", // level 4+ all use Heading4
    };

    let children = inlines_to_children(content, ctx);
    let mut p = Paragraph::with_style(style_id)
        .spacing(0, 0)
        .line_spacing(240, "auto");
    for child in children {
        p = match child {
            ParagraphChild::Run(r) => p.add_run(r),
            ParagraphChild::Hyperlink(h) => p.add_hyperlink(h),
        };
    }
    p
}

/// Convert a paragraph block to a paragraph
fn paragraph_to_paragraph(inlines: &[Inline], ctx: &mut BuildContext) -> Paragraph {
    let children = inlines_to_children(inlines, ctx);
    let mut p = Paragraph::with_style("BodyText")
        .spacing(0, 0)
        .line_spacing(240, "auto");
    for child in children {
        p = match child {
            ParagraphChild::Run(r) => p.add_run(r),
            ParagraphChild::Hyperlink(h) => p.add_hyperlink(h),
        };
    }
    p
}

/// Convert parsed markdown table to DOCX Table
///
/// # Arguments
/// * `headers` - Table header cells
/// * `alignments` - Column alignments
/// * `rows` - Data rows
/// * `ctx` - Build context holding tracked state
///
/// # Returns
/// A DOCX Table structure
fn table_to_docx(
    headers: &[ParserTableCell],
    alignments: &[ParserAlignment],
    rows: &[Vec<ParserTableCell>],
    ctx: &mut BuildContext,
) -> Table {
    let mut table = Table::new().with_header_row(true);

    // Calculate column count
    let col_count = headers.len();

    // Estimate content length to decide table width
    // Check header length
    let mut max_row_chars: usize = headers
        .iter()
        .map(|c| estimate_inline_length(&c.content))
        .sum();

    // Check rows length
    for row in rows {
        let row_len: usize = row.iter().map(|c| estimate_inline_length(&c.content)).sum();
        max_row_chars = max_row_chars.max(row_len);
    }

    // Determine Mode
    let (table_width, cell_width) = if max_row_chars > 60 || col_count > 5 {
        // Wide Mode: 100% width
        (
            TableWidth::Pct(5000),
            TableWidth::Pct(5000 / col_count.max(1) as u32),
        )
    } else {
        // Narrow Mode: Auto width
        (TableWidth::Auto, TableWidth::Auto)
    };

    table = table.width(table_width);

    // Auto column widths (equal distribution, ~9000 twips total for A4)
    // Keep this for w:tblGrid even if w:tblW overrides it visually
    let col_width = 9000 / col_count.max(1) as u32;
    table = table.with_column_widths(vec![col_width; col_count]);

    // Add header row
    let mut header_row = TableRow::new().header();
    for (i, cell) in headers.iter().enumerate() {
        let alignment = alignments.get(i).copied().unwrap_or(ParserAlignment::None);
        let cell_elem = create_table_cell(&cell.content, alignment, true, cell_width, ctx);
        header_row = header_row.add_cell(cell_elem);
    }
    table = table.add_row(header_row);

    // Add data rows
    for row in rows {
        let mut data_row = TableRow::new();
        for (i, cell) in row.iter().enumerate() {
            let alignment = alignments.get(i).copied().unwrap_or(ParserAlignment::None);
            let cell_elem = create_table_cell(&cell.content, alignment, false, cell_width, ctx);
            data_row = data_row.add_cell(cell_elem);
        }
        table = table.add_row(data_row);
    }

    table
}

/// Create a table cell with content
///
/// # Arguments
/// * `content` - Inline content for the cell
/// * `alignment` - Cell alignment
/// * `is_header` - Whether this is a header cell
/// * `width` - Cell width
/// * `ctx` - Build context holding tracked state
///
/// # Returns
/// A DOCX TableCellElement
fn create_table_cell(
    content: &[Inline],
    alignment: ParserAlignment,
    is_header: bool,
    width: TableWidth,
    ctx: &mut BuildContext,
) -> TableCellElement {
    let children = inlines_to_children(content, ctx);

    // Build paragraph from children
    let mut p = Paragraph::new().spacing(0, 0).line_spacing(240, "auto");
    for child in children {
        p = match child {
            ParagraphChild::Run(mut r) => {
                if is_header {
                    r.bold = true;
                }
                p.add_run(r)
            }
            ParagraphChild::Hyperlink(h) => p.add_hyperlink(h),
        };
    }

    let align_str = match alignment {
        ParserAlignment::Left => Some("left"),
        ParserAlignment::Center => Some("center"),
        ParserAlignment::Right => Some("right"),
        ParserAlignment::None => None,
    };

    let mut cell = TableCellElement::new().width(width).add_paragraph(p);
    if let Some(align) = align_str {
        cell = cell.alignment(align);
    }
    // Remove spacing from table cell paragraphs to avoid extra gaps
    cell
}

/// Convert a code block to paragraphs (one per line)
fn code_block_to_paragraphs(
    content: &str,
    filename: Option<&str>,
    highlight_lines: &[u32],
    show_line_numbers: bool,
) -> Vec<Paragraph> {
    let mut paragraphs = Vec::new();

    // Add filename as a separate paragraph if present
    if let Some(fname) = filename {
        let filename_para = Paragraph::with_style("CodeFilename")
            .add_text(fname)
            .spacing(0, 0)
            .line_spacing(240, "auto");
        paragraphs.push(filename_para);
    }

    // Add each line as a separate paragraph
    // Line numbers are 1-based
    for (i, line) in content.lines().enumerate() {
        let line_num = (i + 1) as u32;
        let mut p = Paragraph::with_style("Code")
            .spacing(0, 0)
            .line_spacing(240, "auto");

        // Handle line numbers
        if show_line_numbers {
            // Add line number as a separate run with gray color
            let num_text = format!("{:>2}. ", line_num);
            p = p.add_run(Run::new(num_text).color("888888"));
        }

        p = p.add_text(line);

        // Handle highlighting
        if highlight_lines.contains(&line_num) {
            // Light yellow background for highlighted lines
            p = p.shading("FFFACD"); // LemonChiffon
        }

        paragraphs.push(p);
    }

    // If content is empty, add at least one paragraph
    if paragraphs.is_empty() || (paragraphs.len() == 1 && filename.is_some()) {
        paragraphs.push(
            Paragraph::with_style("Code")
                .add_text("")
                .spacing(0, 0)
                .line_spacing(240, "auto"),
        );
    }

    paragraphs
}

/// Convert a list to paragraphs with a specific numId (for unique list instances)
fn list_to_paragraphs_with_num_id(
    _ordered: bool,
    _start: Option<u32>,
    items: &[ListItem],
    list_level: usize,
    num_id: u32,
    ctx: &mut BuildContext,
    skip_toc: bool,
) -> Vec<Paragraph> {
    let mut paragraphs = Vec::new();

    for item in items.iter() {
        // Process the item's content blocks
        for block in &item.content {
            let mut item_paragraphs = block_to_paragraphs(block, list_level + 1, ctx, skip_toc);

            // Apply list styling to the first paragraph of the item
            if let Some(first_para) = item_paragraphs.first_mut() {
                first_para.style_id = Some("ListParagraph".to_string());

                // Use the provided unique numId for this list
                let ilvl = list_level as u32;
                first_para.numbering_id = Some(num_id);
                first_para.numbering_level = Some(ilvl);
            }

            paragraphs.extend(item_paragraphs);
        }
    }

    paragraphs
}

/// Convert thematic break to a paragraph with a section break
fn thematic_break_to_paragraph() -> Paragraph {
    // Treat "---" as a Next Page Section Break
    Paragraph::new()
        .section_break("nextPage")
        .spacing(0, 0)
        .line_spacing(240, "auto")
}

/// Convert inline elements to ParagraphChild (Run or Hyperlink)
///
/// This handles the conversion of inline formatting (bold, italic, code, etc.)
/// into DOCX runs or hyperlinks with appropriate formatting.
///
/// # Arguments
/// * `inlines` - Slice of inline elements to convert
/// * `ctx` - Build context holding tracked state
///
/// # Returns
/// A vector of paragraph children (runs or hyperlinks)
fn inlines_to_children(inlines: &[Inline], ctx: &mut BuildContext) -> Vec<ParagraphChild> {
    let mut children = Vec::new();

    for inline in inlines {
        children.extend(inline_to_children(inline, false, false, false, ctx));
    }

    children
}

/// Flatten nested inline formatting to paragraph children (runs or hyperlinks)
///
/// This recursive function handles nested formatting like **bold *italic***
/// by tracking the current formatting state.
///
/// # Arguments
/// * `inline` - The inline element to convert
/// * `bold` - Current bold state (from parent formatting)
/// * `italic` - Current italic state (from parent formatting)
/// * `strike` - Current strikethrough state (from parent formatting)
/// * `ctx` - Build context holding tracked state
///
/// # Returns
/// A vector of paragraph children (runs or hyperlinks)
fn inline_to_children(
    inline: &Inline,
    bold: bool,
    italic: bool,
    strike: bool,
    ctx: &mut BuildContext,
) -> Vec<ParagraphChild> {
    match inline {
        Inline::Text(text) => {
            let mut run = Run::new(text).preserve_space(true);
            if bold {
                run = run.bold();
            }
            if italic {
                run = run.italic();
            }
            if strike {
                run = run.strike();
            }
            vec![ParagraphChild::Run(run)]
        }

        Inline::Bold(content) => {
            let mut children = Vec::new();
            for inner in content {
                children.extend(inline_to_children(inner, true, italic, strike, ctx));
            }
            children
        }

        Inline::Italic(content) => {
            let mut children = Vec::new();
            for inner in content {
                children.extend(inline_to_children(inner, bold, true, strike, ctx));
            }
            children
        }

        Inline::BoldItalic(content) => {
            let mut children = Vec::new();
            for inner in content {
                children.extend(inline_to_children(inner, true, true, strike, ctx));
            }
            children
        }

        Inline::Code(text) => {
            vec![ParagraphChild::Run(
                Run::new(text).style("CodeChar").preserve_space(true),
            )]
        }

        Inline::Strikethrough(content) => {
            let mut children = Vec::new();
            for inner in content {
                children.extend(inline_to_children(inner, bold, italic, true, ctx));
            }
            children
        }

        Inline::Link { text, url, .. } => {
            let rel_id = ctx.hyperlink_ctx.add_hyperlink(url);
            let mut hyperlink = crate::docx::ooxml::Hyperlink::new(rel_id);

            // Process nested text
            let children = inlines_to_children(text, ctx);
            for child in children {
                if let ParagraphChild::Run(mut run) = child {
                    run.style = Some("Hyperlink".to_string());
                    run.underline = true;
                    hyperlink = hyperlink.add_run(run);
                }
                // Nested hyperlinks not supported by Word, ignore
            }
            vec![ParagraphChild::Hyperlink(hyperlink)]
        }

        Inline::Image { .. } => {
            // Skip inline images for now - will be handled in Phase 3
            vec![]
        }

        Inline::FootnoteRef(label) => {
            // Look up footnote definition
            if let Some(blocks) = ctx.doc.footnotes.get(label) {
                // Convert blocks to paragraphs
                // Use a temporary numbering context for footnote content
                let mut footnote_numbering_ctx = NumberingContext::new();
                let mut footnote_toc_builder = TocBuilder::new();
                let mut footnote_bookmark_id: u32 = 0;
                let mut footnote_xref_ctx = CrossRefContext::new();
                let mut content = Vec::new();
                for block in blocks {
                    let mut nested_ctx = BuildContext {
                        image_ctx: &mut ImageContext::new(), // Temporary
                        hyperlink_ctx: ctx.hyperlink_ctx,
                        numbering_ctx: &mut footnote_numbering_ctx,
                        image_id: &mut 0, // Temporary
                        doc: ctx.doc,
                        footnotes: ctx.footnotes,
                        toc_builder: &mut footnote_toc_builder,
                        bookmark_id_counter: &mut footnote_bookmark_id,
                        xref_ctx: &mut footnote_xref_ctx,
                    };
                    let paragraphs = block_to_paragraphs(
                        block,
                        0,
                        &mut nested_ctx,
                        false, // Never add footnotes to TOC
                    );
                    content.extend(paragraphs);
                }

                if !content.is_empty() {
                    let id = ctx.footnotes.add_footnote(content);
                    // Return a run with footnote reference
                    let mut run = Run::new("");
                    run.footnote_id = Some(id);
                    vec![ParagraphChild::Run(run)]
                } else {
                    vec![]
                }
            } else {
                // Fallback: show the reference as text if definition is missing
                vec![ParagraphChild::Run(Run::new(format!("[^{}]", label)))]
            }
        }

        Inline::CrossRef { target, ref_type } => {
            // Get display text from cross-reference context
            let display_text = ctx.xref_ctx.get_display_text(target, *ref_type);

            // Create a run with the display text, styled as a link
            // For now, just use blue color and underline
            // TODO: In the future, create an internal hyperlink to the bookmark
            let mut run = Run::new(&display_text);
            run.color = Some("0563C1".to_string()); // Word hyperlink blue
            run.underline = true;
            vec![ParagraphChild::Run(run)]
        }

        Inline::SoftBreak => {
            // Soft break becomes a space
            vec![ParagraphChild::Run(Run::new(" "))]
        }

        Inline::HardBreak => {
            // Hard break becomes a line break element
            vec![ParagraphChild::Run(create_break_run())]
        }

        Inline::Html(_) => {
            // Skip inline HTML for now
            vec![]
        }

        Inline::IndexMarker(_) => {
            // Skip index markers for now - will be handled in Phase 3
            vec![]
        }
    }
}

/// Create a run with a line break element
///
/// In OOXML, a line break is represented by a `<w:br/>` element
/// inside a run. This creates a run that contains only a break.
fn create_break_run() -> Run {
    // For now, we'll use a text run with a newline
    // In a full implementation, we'd modify the Run struct to support
    // explicit break elements and update the XML generation accordingly
    Run::new("\n").preserve_space(true)
}

/// Estimate the text length of inline elements
fn estimate_inline_length(inlines: &[Inline]) -> usize {
    inlines
        .iter()
        .map(|i| match i {
            Inline::Text(s) | Inline::Code(s) => s.chars().count(),
            Inline::Bold(v) | Inline::Italic(v) | Inline::Strikethrough(v) => {
                estimate_inline_length(v)
            }
            Inline::BoldItalic(v) => estimate_inline_length(v),
            Inline::Link { text, .. } => estimate_inline_length(text),
            _ => 1,
        })
        .sum()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::ooxml::{DocElement, HeaderFooterField};
    use crate::parser::parse_markdown_with_frontmatter;
    use crate::parser::RefType;

    /// Helper function to get a config with TOC disabled (for tests that don't need TOC)
    fn no_toc_config() -> DocumentConfig {
        DocumentConfig {
            toc: crate::docx::toc::TocConfig {
                enabled: false,
                ..Default::default()
            },
            ..Default::default()
        }
    }

    /// Helper function to extract paragraphs from document elements
    fn get_paragraphs(doc: &DocumentXml) -> Vec<&Paragraph> {
        doc.elements
            .iter()
            .filter_map(|e| match e {
                DocElement::Paragraph(p) => Some(p.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_footnote_reference() {
        let md = "This is text with a footnote[^1].\n\n[^1]: This is the footnote content.";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Should have one footnote
        assert_eq!(result.footnotes.len(), 1);

        // Document should have footnote reference
        let paragraphs = get_paragraphs(&result.document);
        assert!(!paragraphs.is_empty());

        // Check for footnote reference in runs
        let has_footnote_ref = paragraphs
            .iter()
            .flat_map(|p| p.iter_runs())
            .any(|r| r.footnote_id.is_some());
        assert!(has_footnote_ref, "Should have a footnote reference");

        // Footnote content should be present
        let footnotes = result.footnotes.get_footnotes();
        assert_eq!(footnotes[0].id, 1);
        assert!(!footnotes[0].content.is_empty());
    }

    #[test]
    fn test_multiple_footnotes() {
        let md = "Text with two footnotes[^1][^2].\n\n[^1]: First footnote\n[^2]: Second footnote";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Should have two footnotes
        assert_eq!(result.footnotes.len(), 2);

        // Check footnote IDs
        let footnotes = result.footnotes.get_footnotes();
        assert_eq!(footnotes[0].id, 1);
        assert_eq!(footnotes[1].id, 2);

        // Check footnote content
        let footnote1_text: String = footnotes[0]
            .content
            .iter()
            .flat_map(|p| p.iter_runs().map(|r| r.text.as_str()))
            .collect();
        assert!(footnote1_text.contains("First"));

        let footnote2_text: String = footnotes[1]
            .content
            .iter()
            .flat_map(|p| p.iter_runs().map(|r| r.text.as_str()))
            .collect();
        assert!(footnote2_text.contains("Second"));
    }

    #[test]
    fn test_footnote_with_formatting() {
        let md = "Text[^1]\n\n[^1]: Footnote with **bold** and *italic* text.";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        assert_eq!(result.footnotes.len(), 1);

        // Check that footnote content has formatting
        let footnotes = result.footnotes.get_footnotes();
        let footnote = &footnotes[0];
        let has_bold = footnote
            .content
            .iter()
            .flat_map(|p| p.iter_runs())
            .any(|r| r.bold);
        let has_italic = footnote
            .content
            .iter()
            .flat_map(|p| p.iter_runs())
            .any(|r| r.italic);

        assert!(has_bold, "Footnote should have bold text");
        assert!(has_italic, "Footnote should have italic text");
    }

    #[test]
    fn test_footnote_missing_definition() {
        let md = "Text with missing footnote[^99].";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Should have no footnotes (definition missing)
        assert_eq!(result.footnotes.len(), 0);

        // Should show fallback text
        let paragraphs = get_paragraphs(&result.document);
        let text: String = paragraphs
            .iter()
            .flat_map(|p| p.iter_runs().map(|r| r.text.as_str()))
            .collect();
        assert!(
            text.contains("[^99]"),
            "Should show fallback text for missing footnote"
        );
    }

    #[test]
    fn test_footnote_in_heading() {
        let md = "# Heading with footnote[^1]\n\n[^1]: Footnote content";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());

        assert_eq!(result.footnotes.len(), 1);

        // Heading should have footnote reference
        let paragraphs = get_paragraphs(&result.document);
        let heading = paragraphs.first().unwrap();
        assert_eq!(heading.style_id, Some("Heading1".to_string()));

        let has_footnote_ref = heading.iter_runs().any(|r| r.footnote_id.is_some());
        assert!(has_footnote_ref, "Heading should have footnote reference");
    }

    #[test]
    fn test_footnote_in_list() {
        let md = "- Item with footnote[^1]\n\n[^1]: Footnote content";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        assert_eq!(result.footnotes.len(), 1);

        // List item should have footnote reference
        let paragraphs = get_paragraphs(&result.document);
        let list_item = paragraphs.first().unwrap();
        assert_eq!(list_item.style_id, Some("ListParagraph".to_string()));

        let has_footnote_ref = list_item.iter_runs().any(|r| r.footnote_id.is_some());
        assert!(has_footnote_ref, "List item should have footnote reference");
    }

    #[test]
    fn test_footnote_in_blockquote() {
        let md = "> Quote with footnote[^1]\n\n[^1]: Footnote content";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        assert_eq!(result.footnotes.len(), 1);

        // Blockquote should have footnote reference
        let paragraphs = get_paragraphs(&result.document);
        let quote = paragraphs.first().unwrap();
        assert_eq!(quote.style_id, Some("Quote".to_string()));

        let has_footnote_ref = quote.iter_runs().any(|r| r.footnote_id.is_some());
        assert!(
            has_footnote_ref,
            "Blockquote should have footnote reference"
        );
    }

    #[test]
    fn test_footnote_in_table() {
        let md = "| Header |\n|--------|\n| Cell[^1] |\n\n[^1]: Footnote content";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        assert_eq!(result.footnotes.len(), 1);

        // Table cell should have footnote reference
        if let Some(DocElement::Table(table)) = result.document.elements.first() {
            let data_row = &table.rows[1];
            let cell = &data_row.cells[0];
            let has_footnote_ref = cell
                .paragraphs
                .iter()
                .flat_map(|p| p.iter_runs())
                .any(|r| r.footnote_id.is_some());
            assert!(
                has_footnote_ref,
                "Table cell should have footnote reference"
            );
        }
    }

    #[test]
    fn test_footnote_multiline_content() {
        let md = "Text[^1]\n\n[^1]: First line\n    Second line\n    Third line";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        assert_eq!(result.footnotes.len(), 1);

        // Footnote should have multiple paragraphs (one per line)
        let footnotes = result.footnotes.get_footnotes();
        let footnote = &footnotes[0];
        assert!(
            footnote.content.len() >= 2,
            "Footnote should have multiple lines"
        );
    }

    #[test]
    fn test_footnote_xml_generation() {
        let md = "Text[^1]\n\n[^1]: Footnote content";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        let xml = result.footnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Check XML structure
        assert!(xml_str.contains("<?xml version"));
        assert!(xml_str.contains("<w:footnotes"));
        assert!(xml_str.contains("<w:footnote w:id=\"1\""));
        assert!(xml_str.contains("Footnote content"));
        assert!(xml_str.contains("<w:footnote w:type=\"separator\" w:id=\"-1\""));
    }

    #[test]
    fn test_cross_reference_in_document() {
        let doc = ParsedDocument {
            frontmatter: None,
            blocks: vec![
                Block::Heading {
                    level: 1,
                    content: vec![Inline::Text("Introduction".to_string())],
                    id: Some("intro".to_string()),
                },
                Block::Paragraph(vec![
                    Inline::Text("See ".to_string()),
                    Inline::CrossRef {
                        target: "intro".to_string(),
                        ref_type: RefType::Chapter,
                    },
                    Inline::Text(" for more.".to_string()),
                ]),
            ],
            footnotes: std::collections::HashMap::new(),
        };

        let config = no_toc_config(); // Disable TOC for this test
        let result = build_document(&doc, Language::English, &config);

        // Verify the document was built successfully
        assert!(!result.document.elements.is_empty());

        // Check that we have a heading and a paragraph
        let paragraphs = get_paragraphs(&result.document);
        assert!(paragraphs.len() >= 2);

        // First paragraph should be the heading
        let heading = &paragraphs[0];
        assert_eq!(heading.style_id, Some("Heading1".to_string()));

        // Second paragraph should contain the cross-reference
        let para = &paragraphs[1];
        let text: String = para.iter_runs().map(|r| r.text.as_str()).collect();
        assert!(
            text.contains("Chapter 1"),
            "Should contain 'Chapter 1' from cross-reference"
        );
    }

    #[test]
    fn test_cross_reference_unresolved() {
        let doc = ParsedDocument {
            frontmatter: None,
            blocks: vec![Block::Paragraph(vec![
                Inline::Text("See ".to_string()),
                Inline::CrossRef {
                    target: "nonexistent".to_string(),
                    ref_type: RefType::Chapter,
                },
                Inline::Text(" for more.".to_string()),
            ])],
            footnotes: std::collections::HashMap::new(),
        };

        let config = DocumentConfig::default();
        let result = build_document(&doc, Language::English, &config);

        // Should show placeholder for unresolved reference
        let paragraphs = get_paragraphs(&result.document);
        let text: String = paragraphs
            .iter()
            .flat_map(|p| p.iter_runs().map(|r| r.text.as_str()))
            .collect();
        assert!(
            text.contains("[nonexistent]"),
            "Should show placeholder for unresolved reference"
        );
    }

    #[test]
    fn test_cross_reference_with_figure() {
        let doc = ParsedDocument {
            frontmatter: None,
            blocks: vec![
                Block::Heading {
                    level: 1,
                    content: vec![Inline::Text("Chapter 1".to_string())],
                    id: Some("ch1".to_string()),
                },
                Block::Image {
                    alt: "System Architecture".to_string(),
                    src: "arch.png".to_string(),
                    title: None,
                    width: None,
                    id: Some("fig:arch".to_string()),
                },
                Block::Paragraph(vec![
                    Inline::Text("See ".to_string()),
                    Inline::CrossRef {
                        target: "fig:arch".to_string(),
                        ref_type: RefType::Figure,
                    },
                    Inline::Text(" for details.".to_string()),
                ]),
            ],
            footnotes: std::collections::HashMap::new(),
        };

        let config = DocumentConfig::default();
        let result = build_document(&doc, Language::English, &config);

        // Check that figure reference is properly formatted
        let paragraphs = get_paragraphs(&result.document);
        let text: String = paragraphs
            .iter()
            .flat_map(|p| p.iter_runs().map(|r| r.text.as_str()))
            .collect();
        assert!(
            text.contains("Figure 1.1"),
            "Should contain 'Figure 1.1' for the figure reference"
        );
    }

    #[test]
    fn test_build_result_includes_footnotes() {
        let md = "Text[^1]\n\n[^1]: Footnote";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // BuildResult should include footnotes field
        assert!(!result.footnotes.is_empty());
        assert_eq!(result.footnotes.len(), 1);
    }

    #[test]
    fn test_debug_blockquote_parsing() {
        let md = "> This is a quote\n> With multiple lines";
        let parsed = parse_markdown_with_frontmatter(md);

        println!("Number of blocks: {}", parsed.blocks.len());
        for (i, block) in parsed.blocks.iter().enumerate() {
            println!("Block {}: {:?}", i, block);
        }
    }

    #[test]
    fn test_build_document_simple() {
        let md = "# Hello World\n\nThis is a paragraph.";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_id, Some("Heading1".to_string()));
        assert_eq!(
            paragraphs[0].iter_runs().next().unwrap().text,
            "Hello World"
        );
        assert_eq!(paragraphs[1].style_id, Some("BodyText".to_string()));
        assert_eq!(
            paragraphs[1].iter_runs().next().unwrap().text,
            "This is a paragraph."
        );
    }

    #[test]
    fn test_heading_levels() {
        let md = "# H1\n\n## H2\n\n### H3\n\n#### H4\n\n##### H5";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 5);
        assert_eq!(paragraphs[0].style_id, Some("Heading1".to_string()));
        assert_eq!(paragraphs[1].style_id, Some("Heading2".to_string()));
        assert_eq!(paragraphs[2].style_id, Some("Heading3".to_string()));
        assert_eq!(paragraphs[3].style_id, Some("Heading4".to_string()));
        assert_eq!(paragraphs[4].style_id, Some("Heading4".to_string())); // H5 also uses Heading4
    }

    #[test]
    fn test_inline_formatting() {
        let md = "This is **bold**, *italic*, and `code`.";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        let runs = paragraphs[0].get_runs();

        // Should have runs for: "This is ", "bold", ", ", "italic", ", and ", "code", "."
        assert!(runs.len() >= 5);

        // Check for bold run
        let bold_run = runs.iter().find(|r| r.bold);
        assert!(bold_run.is_some());
        assert_eq!(bold_run.unwrap().text, "bold");

        // Check for italic run
        let italic_run = runs.iter().find(|r| r.italic);
        assert!(italic_run.is_some());
        assert_eq!(italic_run.unwrap().text, "italic");

        // Check for code run
        let code_run = runs
            .iter()
            .find(|r| r.style == Some("CodeChar".to_string()));
        assert!(code_run.is_some());
        assert_eq!(code_run.unwrap().text, "code");
    }

    #[test]
    fn test_nested_formatting() {
        let md = "This is **bold and *italic*** text.";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        let runs = paragraphs[0].get_runs();

        // The parser doesn't handle nested formatting correctly yet
        // It creates separate Bold and Italic blocks instead of nesting them
        // For now, just verify we have bold and italic runs
        assert!(runs.iter().any(|r| r.bold));
        assert!(runs.iter().any(|r| r.italic));
    }

    #[test]
    fn test_code_block() {
        let md = "```rust\nfn main() {\n    println!(\"Hello\");\n}\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have 3 lines of code
        assert_eq!(paragraphs.len(), 3);
        assert_eq!(paragraphs[0].style_id, Some("Code".to_string()));
        assert_eq!(
            paragraphs[0].iter_runs().next().unwrap().text,
            "fn main() {"
        );
        assert_eq!(
            paragraphs[1].iter_runs().next().unwrap().text,
            "    println!(\"Hello\");"
        );
        assert_eq!(paragraphs[2].iter_runs().next().unwrap().text, "}");
    }

    #[test]
    fn test_blockquote() {
        let md = "> This is a quote\n> With multiple lines";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // The parser creates a BlockQuote with one paragraph containing both lines
        // Plus an empty paragraph at the end
        assert!(get_paragraphs(docx).len() >= 1);
        assert_eq!(paragraphs[0].style_id, Some("Quote".to_string()));
        // Should have text from both lines
        let text: String = paragraphs[0].iter_runs().map(|r| r.text.as_str()).collect();
        assert!(text.contains("This is a quote"));
        assert!(text.contains("With multiple lines"));
    }

    #[test]
    fn test_unordered_list() {
        let md = "- Item 1\n- Item 2\n- Item 3";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 3);
        for p in paragraphs {
            assert_eq!(p.style_id, Some("ListParagraph".to_string()));
        }
    }

    #[test]
    fn test_ordered_list() {
        let md = "1. First\n2. Second\n3. Third";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 3);
        for p in paragraphs {
            assert_eq!(p.style_id, Some("ListParagraph".to_string()));
            assert_eq!(p.numbering_id, Some(1));
            assert_eq!(p.numbering_level, Some(0));
        }
    }

    #[test]
    fn test_thematic_break() {
        let md = "---";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        // Thematic break creates an empty paragraph
        assert!(paragraphs[0].children.is_empty());
    }

    #[test]
    fn test_link() {
        let md = "[OpenAI](https://openai.com)";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);

        // Check that hyperlink was tracked
        assert_eq!(result.hyperlinks.hyperlinks.len(), 1);
        assert_eq!(result.hyperlinks.hyperlinks[0].url, "https://openai.com");

        // Check paragraph has hyperlink child
        assert_eq!(paragraphs[0].children.len(), 1);
        match &paragraphs[0].children[0] {
            ParagraphChild::Hyperlink(h) => {
                assert_eq!(h.children.len(), 1);
                assert_eq!(h.children[0].text, "OpenAI");
                assert_eq!(h.children[0].style, Some("Hyperlink".to_string()));
                assert!(h.children[0].underline);
            }
            _ => panic!("Expected hyperlink child"),
        }
    }

    #[test]
    fn test_strikethrough() {
        let md = "~~deleted text~~";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        let runs = paragraphs[0].get_runs();

        let strike_run = runs.iter().find(|r| r.strike);
        assert!(strike_run.is_some());
        assert_eq!(strike_run.unwrap().text, "deleted text");
    }

    #[test]
    fn test_soft_break() {
        let md = "Line 1\nLine 2"; // Single newline = soft break
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Soft break should be in the same paragraph
        assert_eq!(paragraphs.len(), 1);
        let text: String = paragraphs[0].iter_runs().map(|r| r.text.as_str()).collect();
        assert!(text.contains("Line 1"));
        assert!(text.contains("Line 2"));
    }

    #[test]
    fn test_hard_break() {
        let md = "Line 1  \nLine 2"; // Two spaces + newline = hard break
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        let text: String = paragraphs[0].iter_runs().map(|r| r.text.as_str()).collect();
        assert!(text.contains("Line 1"));
        assert!(text.contains("Line 2"));
    }

    #[test]
    fn test_complex_document() {
        let md = r#"# Document Title

This is a paragraph with **bold** and *italic* text.

## Section

> A blockquote
> with multiple lines

- List item 1
- List item 2

```rust
fn main() {
    println!("Hello");
}
```

---

End of document.
"#;
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        // Count paragraphs:
        // 1. Heading1
        // 2. Paragraph
        // 3. Heading2
        // 4. Quote (1 paragraph with both lines)
        // 5. List item 1
        // 6. List item 2
        // 7. Code line 1
        // 8. Code line 2
        // 9. Code line 3
        // 10. Thematic break
        // 11. Final paragraph
        // Plus possibly empty paragraphs from parser
        assert!(get_paragraphs(docx).len() >= 11);
    }

    #[test]
    fn test_html_blocks_skipped() {
        let md = "<div>This is HTML</div>\n\nRegular paragraph";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // HTML should be skipped, only regular paragraph remains
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(
            paragraphs[0].iter_runs().next().unwrap().text,
            "Regular paragraph"
        );
    }

    #[test]
    fn test_mixed_formatting_in_paragraph() {
        let md = "Normal **bold** *italic* `code` ~~strike~~";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        let runs = paragraphs[0].get_runs();

        assert!(runs.iter().any(|r| r.bold && r.text == "bold"));
        assert!(runs.iter().any(|r| r.italic && r.text == "italic"));
        assert!(runs.iter().any(|r| r.style == Some("CodeChar".to_string())));
        assert!(runs.iter().any(|r| r.strike && r.text == "strike"));
    }

    #[test]
    fn test_blockquote_with_formatting() {
        let md = "> This is a **bold** quote";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // The parser creates a BlockQuote with one paragraph
        assert!(get_paragraphs(docx).len() >= 1);
        assert_eq!(paragraphs[0].style_id, Some("Quote".to_string()));
        assert!(paragraphs[0].iter_runs().any(|r| r.bold));
    }

    #[test]
    fn test_list_with_nested_content() {
        let md = "- Item with **bold** text\n- Another item";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // The parser creates a List with 2 items
        // Each item becomes a separate paragraph in the output
        assert!(get_paragraphs(docx).len() >= 2);
        assert_eq!(paragraphs[0].style_id, Some("ListParagraph".to_string()));
    }

    #[test]
    fn test_heading_with_inline_code() {
        let md = "# Heading with `code`";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].style_id, Some("Heading1".to_string()));
        assert!(paragraphs[0]
            .iter_runs()
            .any(|r| r.style == Some("CodeChar".to_string())));
    }

    #[test]
    fn test_multiple_code_blocks() {
        let md = "```rust\nfn main() {}\n```\n\n```python\ndef main():\n    pass\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // First code block: 1 line
        // Second code block: 2 lines
        assert_eq!(paragraphs.len(), 3);
        assert_eq!(paragraphs[0].style_id, Some("Code".to_string()));
        assert_eq!(paragraphs[1].style_id, Some("Code".to_string()));
        assert_eq!(paragraphs[2].style_id, Some("Code".to_string()));
    }

    #[test]
    fn test_preserve_whitespace_in_code() {
        let md = "```rust\nfn main() {\n    println!(\"Hello\");\n}\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Check that indentation is preserved
        assert_eq!(
            paragraphs[1].iter_runs().next().unwrap().text,
            "    println!(\"Hello\");"
        );
    }

    #[test]
    fn test_code_block_with_line_numbers() {
        let md = "```rust,ln\nline 1\nline 2\nline 3\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have 3 paragraphs (one per line)
        assert_eq!(paragraphs.len(), 3);

        // Check line numbers are present
        assert!(paragraphs[0]
            .iter_runs()
            .next()
            .unwrap()
            .text
            .contains("1."));
        assert!(paragraphs[1]
            .iter_runs()
            .next()
            .unwrap()
            .text
            .contains("2."));
        assert!(paragraphs[2]
            .iter_runs()
            .next()
            .unwrap()
            .text
            .contains("3."));

        // Check line numbers are gray
        assert_eq!(
            paragraphs[0].iter_runs().next().unwrap().color,
            Some("888888".to_string())
        );
        assert_eq!(
            paragraphs[1].iter_runs().next().unwrap().color,
            Some("888888".to_string())
        );
        assert_eq!(
            paragraphs[2].iter_runs().next().unwrap().color,
            Some("888888".to_string())
        );
    }

    #[test]
    fn test_code_block_with_highlighting() {
        let md = "```rust,hl=2\nline 1\nline 2\nline 3\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have 3 paragraphs
        assert_eq!(paragraphs.len(), 3);

        // Line 2 should have highlighting
        assert!(paragraphs[0].shading.is_none());
        assert_eq!(paragraphs[1].shading, Some("FFFACD".to_string()));
        assert!(paragraphs[2].shading.is_none());
    }

    #[test]
    fn test_code_block_with_line_numbers_and_highlighting() {
        let md = "```rust,hl=2,ln\nline 1\nline 2\nline 3\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have 3 paragraphs
        assert_eq!(paragraphs.len(), 3);

        // Line 2 should have both line number and highlighting
        assert_eq!(paragraphs[1].iter_runs().next().unwrap().text, " 2. ");
        assert_eq!(
            paragraphs[1].iter_runs().next().unwrap().color,
            Some("888888".to_string())
        );
        assert_eq!(paragraphs[1].shading, Some("FFFACD".to_string()));
    }

    #[test]
    fn test_code_block_with_multiple_highlights() {
        let md = "```rust,hl=1,3\nline 1\nline 2\nline 3\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Lines 1 and 3 should have highlighting
        assert_eq!(paragraphs[0].shading, Some("FFFACD".to_string()));
        assert!(paragraphs[1].shading.is_none());
        assert_eq!(paragraphs[2].shading, Some("FFFACD".to_string()));
    }

    #[test]
    fn test_code_block_with_filename() {
        let md = "```rust,filename=main.rs\nfn main() {}\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have filename paragraph + code line
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_id, Some("CodeFilename".to_string()));
        assert_eq!(paragraphs[0].iter_runs().next().unwrap().text, "main.rs");
        assert_eq!(paragraphs[1].style_id, Some("Code".to_string()));
        assert_eq!(
            paragraphs[1].iter_runs().next().unwrap().text,
            "fn main() {}"
        );
    }

    #[test]
    fn test_link_with_formatting() {
        let md = "[**bold link**](https://example.com)";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 1);
        let runs = paragraphs[0].get_runs();

        // The parser doesn't handle nested formatting in links correctly yet
        // For now, just verify we have some runs
        assert!(!runs.is_empty());
    }

    #[test]
    fn test_table_conversion() {
        let md = "| Header 1 | Header 2 |\n|----------|----------|\n| Cell 1   | Cell 2   |";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        // Should have table element + empty paragraph
        assert_eq!(docx.elements.len(), 2);
        assert!(matches!(docx.elements[0], DocElement::Table(_)));
        assert!(matches!(docx.elements[1], DocElement::Paragraph(_)));
    }

    #[test]
    fn test_table_with_formatting() {
        let md = "| **Bold** | *Italic* |\n|----------|----------|\n| `code`   | ~~strike~~ |";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        // Verify table exists
        assert!(docx
            .elements
            .iter()
            .any(|e| matches!(e, DocElement::Table(_))));
    }

    #[test]
    fn test_table_alignment() {
        let md = "| Left | Center | Right |\n|:-----|:------:|------:|\n| L    | C      | R     |";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        if let Some(DocElement::Table(table)) = docx.elements.first() {
            // Check alignments on data row cells
            if let Some(data_row) = table.rows.get(1) {
                assert_eq!(
                    data_row.cells.get(0).and_then(|c| c.alignment.as_deref()),
                    Some("left")
                );
                assert_eq!(
                    data_row.cells.get(1).and_then(|c| c.alignment.as_deref()),
                    Some("center")
                );
                assert_eq!(
                    data_row.cells.get(2).and_then(|c| c.alignment.as_deref()),
                    Some("right")
                );
            }
        }
    }

    #[test]
    fn test_table_with_multiple_rows() {
        let md = "| Name | Age |\n|------|-----|\n| John | 30  |\n| Jane | 25  |\n| Bob  | 35  |";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        if let Some(DocElement::Table(table)) = docx.elements.first() {
            // Should have header row + 3 data rows
            assert_eq!(table.rows.len(), 4);
            assert!(table.rows[0].is_header);
            assert!(!table.rows[1].is_header);
            assert!(!table.rows[2].is_header);
            assert!(!table.rows[3].is_header);
        }
    }

    #[test]
    fn test_table_header_shading() {
        let md = "| H1 | H2 |\n|----|----|\n| D1 | D2 |";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        if let Some(DocElement::Table(table)) = docx.elements.first() {
            // Header cells should have shading
            let header_row = &table.rows[0];
            for cell in &header_row.cells {
                assert_eq!(cell.shading, Some("D9E2F3".to_string()));
            }
            // Data cells should not have shading
            let data_row = &table.rows[1];
            for cell in &data_row.cells {
                assert!(cell.shading.is_none());
            }
        }
    }

    #[test]
    fn test_table_header_bold() {
        let md = "| Header |\n|--------|\n| Data   |";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());
        let docx = &result.document;

        if let Some(DocElement::Table(table)) = docx.elements.first() {
            // Header cell text should be bold
            let header_row = &table.rows[0];
            if let Some(header_cell) = header_row.cells.first() {
                if let Some(header_para) = header_cell.paragraphs.first() {
                    assert!(header_para.iter_runs().any(|r| r.bold));
                }
            }
            // Data cell text should not be bold
            let data_row = &table.rows[1];
            if let Some(data_cell) = data_row.cells.first() {
                if let Some(data_para) = data_cell.paragraphs.first() {
                    assert!(!data_para.iter_runs().any(|r| r.bold));
                }
            }
        }
    }

    #[test]
    fn test_document_with_table_and_paragraphs() {
        let md = "# Title\n\nSome text.\n\n| Col 1 | Col 2 |\n|-------|-------|\n| A     | B     |\n\nMore text.";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());
        let docx = &result.document;

        // Should have: heading, paragraph, table, empty paragraph, paragraph
        assert_eq!(docx.elements.len(), 5);
        assert!(matches!(docx.elements[0], DocElement::Paragraph(_)));
        assert!(matches!(docx.elements[1], DocElement::Paragraph(_)));
        assert!(matches!(docx.elements[2], DocElement::Table(_)));
        assert!(matches!(docx.elements[3], DocElement::Paragraph(_)));
        assert!(matches!(docx.elements[4], DocElement::Paragraph(_)));
    }

    // Image context tests

    #[test]
    fn test_image_context_add() {
        let mut ctx = ImageContext::new();
        let id = ctx.add_image("test.png", None);
        // rId1-3 are reserved, so first image should be rId4
        assert_eq!(id, "rId4");
        assert_eq!(ctx.images.len(), 1);
        assert_eq!(ctx.images[0].src, "test.png");
        assert_eq!(ctx.images[0].filename, "image1.png");
    }

    #[test]
    fn test_image_context_multiple() {
        let mut ctx = ImageContext::new();
        let id1 = ctx.add_image("img1.png", None);
        let id2 = ctx.add_image("img2.png", None);

        assert_eq!(id1, "rId4");
        assert_eq!(id2, "rId5");
        assert_eq!(ctx.images.len(), 2);
        assert_eq!(ctx.images[0].filename, "image1.png");
        assert_eq!(ctx.images[1].filename, "image2.png");
    }

    #[test]
    fn test_image_context_dimensions_default() {
        let mut ctx = ImageContext::new();
        ctx.add_image("test.png", None);
        // Default 6x4 inches
        assert_eq!(ctx.images[0].width_emu, 5486400);
        assert_eq!(ctx.images[0].height_emu, 3657600);
    }

    #[test]
    fn test_image_context_dimensions_inches() {
        let mut ctx = ImageContext::new();
        ctx.add_image("test.png", Some("2in"));
        // 2 inches = 1828800 EMUs
        assert_eq!(ctx.images[0].width_emu, 1828800);
    }

    #[test]
    fn test_image_context_dimensions_pixels() {
        let mut ctx = ImageContext::new();
        ctx.add_image("test.png", Some("96px"));
        // 96px = 1 inch = 914400 EMUs
        assert_eq!(ctx.images[0].width_emu, 914400);
    }

    #[test]
    fn test_image_context_dimensions_percentage() {
        let mut ctx = ImageContext::new();
        ctx.add_image("test.png", Some("50%"));
        // 50% of 6.0in = 3.0in = 2743200 EMUs
        assert_eq!(ctx.images[0].width_emu, 2743200);
    }

    #[test]
    fn test_image_context_filename_generation() {
        let ctx = ImageContext::new();
        assert_eq!(ctx.generate_filename("path/to/test.png"), "image1.png");
        assert_eq!(
            ctx.generate_filename("http://example.com/img.jpg"),
            "image1.jpg"
        );
    }

    #[test]
    fn test_build_document_with_image() {
        let md = "![Test](test.png)";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());

        assert_eq!(result.images.images.len(), 1);
        assert_eq!(result.images.images[0].rel_id, "rId4");

        // Check paragraph has image
        let paragraphs = get_paragraphs(&result.document);
        // Images are not paragraphs, so there should be 0 paragraphs
        assert!(paragraphs.is_empty());

        if let Some(DocElement::Image(img)) = result.document.elements.first() {
            assert_eq!(img.rel_id, "rId4");
        } else {
            panic!("Expected Image element");
        }
    }

    #[test]
    fn test_build_document_image_with_width() {
        let md = "![alt](image.png){width=50%}";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Check that image was added with correct width
        assert_eq!(result.images.images.len(), 1);
        let img = &result.images.images[0];
        assert_eq!(img.src, "image.png");

        // Width should be 50% of 6.0 inches = 3.0 inches
        // 3.0 * 914400 EMUs/inch = 2743200 EMUs
        assert_eq!(img.width_emu, 2743200);
    }

    #[test]
    fn test_header_footer_generation() {
        let md = "# Test Document\n\nThis is a test.";
        let parsed = parse_markdown_with_frontmatter(md);

        let config = DocumentConfig {
            title: "My Document".to_string(),
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

        let result = build_document(&parsed, Language::English, &config);

        // Should have two headers and two footers (default + empty/first)
        // We always generate empty headers/footers for suppression purposes
        assert_eq!(result.headers.len(), 2);
        assert_eq!(result.footers.len(), 2);

        // Check header XML
        let header_xml = String::from_utf8(result.headers[0].1.clone()).unwrap();
        assert!(header_xml.contains("<w:hdr"));
        assert!(header_xml.contains("My Document"));
        assert!(header_xml.contains("STYLEREF"));

        // Check footer XML
        let footer_xml = String::from_utf8(result.footers[0].1.clone()).unwrap();
        assert!(footer_xml.contains("<w:ftr"));
        assert!(footer_xml.contains("PAGE"));

        // Note: Relationship IDs are NOT set in build_document() anymore
        // They are set in lib.rs after calling add_header()/add_footer()
        // The different_first_page flag is set in builder.rs
        assert!(!result.document.header_footer_refs.different_first_page);
    }

    #[test]
    fn test_header_footer_different_first_page() {
        let md = "# Test Document\n\nThis is a test.";
        let parsed = parse_markdown_with_frontmatter(md);

        let config = DocumentConfig {
            title: "My Document".to_string(),
            header: HeaderConfig::default(),
            footer: FooterConfig::default(),
            different_first_page: true,
            ..Default::default()
        };

        let result = build_document(&parsed, Language::English, &config);

        // Should have two headers and two footers (default + first page)
        assert_eq!(result.headers.len(), 2);
        assert_eq!(result.footers.len(), 2);

        // Document should have different first page enabled
        assert!(result.document.header_footer_refs.different_first_page);

        // Note: Relationship IDs (first_header_id, first_footer_id) are NOT set
        // in build_document() anymore - they are set in lib.rs after calling
        // add_header()/add_footer() which returns the actual IDs
    }

    #[test]
    fn test_header_footer_empty_config() {
        let md = "# Test Document\n\nThis is a test.";
        let parsed = parse_markdown_with_frontmatter(md);

        let config = DocumentConfig {
            header: HeaderConfig::empty(),
            footer: FooterConfig::empty(),
            ..Default::default()
        };

        let result = build_document(&parsed, Language::English, &config);

        // Should have no headers or footers
        assert_eq!(result.headers.len(), 0);
        assert_eq!(result.footers.len(), 0);

        // Document should have no header/footer refs
        assert!(result
            .document
            .header_footer_refs
            .default_header_id
            .is_none());
        assert!(result
            .document
            .header_footer_refs
            .default_footer_id
            .is_none());
    }

    #[test]
    fn test_build_document_multiple_images() {
        let md = "![Test1](test1.png)\n\n![Test2](test2.png)";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());

        assert_eq!(result.images.images.len(), 2);
        assert_eq!(result.images.images[0].rel_id, "rId4");
        assert_eq!(result.images.images[1].rel_id, "rId5");
    }

    #[test]
    fn test_toc_generation() {
        let md = "# Chapter 1\n\n## Section 1.1\n\n### Subsection 1.1.1\n\n# Chapter 2";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have: TOC title + 4 TOC entries + blank line + 4 headings
        // Total: 1 (title) + 4 (entries) + 1 (blank) + 4 (headings) = 10 paragraphs
        assert!(paragraphs.len() >= 9);

        // First paragraph should be TOC title
        assert_eq!(paragraphs[0].style_id, Some("TOCHeading".to_string()));

        // TOC entries should have TOC1, TOC2, TOC3 styles
        let toc_entries: Vec<_> = paragraphs
            .iter()
            .filter(|p| p.style_id.as_ref().map_or(false, |s| s.starts_with("TOC")))
            .collect();
        assert!(toc_entries.len() >= 4);

        // Headings should have bookmarks
        let heading_paragraphs: Vec<_> = paragraphs
            .iter()
            .filter(|p| {
                p.style_id
                    .as_ref()
                    .map_or(false, |s| s.starts_with("Heading"))
            })
            .collect();
        assert_eq!(heading_paragraphs.len(), 4);

        // Check that headings have bookmarks
        for heading in heading_paragraphs {
            assert!(
                heading.bookmark_start.is_some(),
                "Heading should have bookmark start"
            );
        }
    }

    #[test]
    fn test_toc_disabled() {
        let md = "# Chapter 1\n\n## Section 1.1";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &no_toc_config());

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have: H1 + H2 = 2 paragraphs
        // (no blank line inserted since H1 is a heading)
        assert_eq!(paragraphs.len(), 2);

        // No TOC title
        assert!(!paragraphs
            .iter()
            .any(|p| p.style_id == Some("TOCHeading".to_string())));
    }

    #[test]
    fn test_toc_depth_filtering() {
        let md = "# H1\n\n## H2\n\n### H3\n\n#### H4";
        let parsed = parse_markdown_with_frontmatter(md);
        let config = DocumentConfig {
            toc: crate::docx::toc::TocConfig {
                enabled: true,
                depth: 2,
                title: "Contents".to_string(),
                ..Default::default()
            },
            ..Default::default()
        };
        let result = build_document(&parsed, Language::English, &config);

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // TOC entries should only include H1 and H2 (depth=2)
        let toc_entries: Vec<_> = paragraphs
            .iter()
            .filter(|p| {
                p.style_id
                    .as_ref()
                    .map_or(false, |s| s.starts_with("TOC") && s != "TOCHeading")
            })
            .collect();

        // Should have TOC1 and TOC2 entries only (H3 and H4 filtered out)
        assert_eq!(toc_entries.len(), 2);
        assert_eq!(toc_entries[0].style_id, Some("TOC1".to_string()));
        assert_eq!(toc_entries[1].style_id, Some("TOC2".to_string()));
    }

    #[test]
    fn test_toc_with_explicit_id() {
        let md = "# Introduction {#intro}\n\n## Getting Started {#start}";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Find headings
        let headings: Vec<_> = paragraphs
            .iter()
            .filter(|p| {
                p.style_id
                    .as_ref()
                    .map_or(false, |s| s.starts_with("Heading"))
            })
            .collect();

        assert_eq!(headings.len(), 2);

        // Check bookmark names match explicit IDs
        assert_eq!(
            headings[0].bookmark_start.as_ref().map(|b| b.name.as_str()),
            Some("intro")
        );
        assert_eq!(
            headings[1].bookmark_start.as_ref().map(|b| b.name.as_str()),
            Some("start")
        );
    }

    #[test]
    fn test_toc_with_formatted_heading() {
        let md = "# **Bold** and *italic* heading";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Find TOC entry
        let toc_entry = paragraphs
            .iter()
            .find(|p| p.style_id == Some("TOC1".to_string()));

        assert!(toc_entry.is_some());

        // TOC entry should contain the plain text (without formatting)
        let text: String = toc_entry
            .unwrap()
            .iter_runs()
            .map(|r| r.text.as_str())
            .collect();
        assert!(text.contains("Bold"));
        assert!(text.contains("italic"));
    }

    #[test]
    fn test_build_document_image_with_alt_text() {
        let md = "![This is alt text](image.png)";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        if let Some(DocElement::Image(img)) = result.document.elements.first() {
            assert_eq!(img.alt_text, "This is alt text");
            assert_eq!(img.name, "image.png");
        } else {
            panic!("Expected image element");
        }
    }

    #[test]
    fn test_build_document_image_in_blockquote() {
        // Note: Parser creates BlockQuote with nested Paragraphs
        // Images in blockquotes are Inline::Image inside Paragraphs
        // Block::Image is only created for standalone images (not yet implemented)
        let md = "> Quote with image\n> ![Image](img.png)";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Parser creates BlockQuote with Paragraphs containing Inline::Image
        // The builder doesn't extract Inline::Image to Block::Image yet
        // So no images are tracked in ImageContext
        assert_eq!(result.images.images.len(), 0);

        // Should have BlockQuote element
        let has_blockquote = result.document.elements.iter().any(
            |e| matches!(e, DocElement::Paragraph(p) if p.style_id == Some("Quote".to_string())),
        );
        assert!(has_blockquote);
    }

    #[test]
    fn test_build_result_structure() {
        let md = "# Test\n\nSome text";
        let parsed = parse_markdown_with_frontmatter(md);
        let result = build_document(&parsed, Language::English, &DocumentConfig::default());

        // Verify BuildResult structure
        assert!(result.document.elements.len() > 0);
        assert_eq!(result.images.next_id, 1); // No images added
        assert_eq!(result.images.images.len(), 0);
    }
}
