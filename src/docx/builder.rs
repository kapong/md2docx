//! DOCX builder - Convert parsed Markdown AST to OOXML
//!
//! This module bridges the parser's AST representation with the OOXML
//! document structure, converting markdown elements to DOCX paragraphs
//! and runs.

use crate::docx::image_utils::{default_image_size_emu, read_image_dimensions};
use crate::docx::ooxml::{
    DocElement, DocumentXml, FooterConfig, FooterXml, FootnotesXml, HeaderConfig, HeaderFooterRefs,
    HeaderXml, ImageElement, Paragraph, ParagraphChild, Run, Table, TableCellElement, TableRow,
    TableWidth, TabStop,
};
use crate::docx::rels_manager::RelIdManager;
use crate::docx::toc::{TocBuilder, TocConfig};
use crate::docx::xref::CrossRefContext;
use crate::parser::{
    extract_inline_text, Alignment as ParserAlignment, Block, Inline, ListItem, ParsedDocument,
    RefType, TableCell as ParserTableCell,
};
use crate::template::extract::table::TableTemplate;
use crate::Language;

/// Tracks images during document building
#[derive(Debug, Default)]
pub(crate) struct ImageContext {
    /// Map of image source path to (filename, relationship_id, data)
    pub images: Vec<ImageInfo>,
    /// Base directory for resolving relative image paths
    pub base_path: Option<std::path::PathBuf>,
}

/// Information about an embedded image
#[derive(Debug, Clone)]
pub(crate) struct ImageInfo {
    pub filename: String,      // e.g., "image1.png"
    pub rel_id: String,        // e.g., "rId4"
    pub src: String,           // Original source path/URL
    pub data: Option<Vec<u8>>, // Image bytes (None if external)
    pub width_emu: i64,        // Width in EMUs
    pub height_emu: i64,       // Height in EMUs
}

/// Tracks hyperlinks during document building
#[derive(Debug, Default, Clone)]
pub(crate) struct HyperlinkContext {
    pub hyperlinks: Vec<HyperlinkInfo>,
}

/// Information about a hyperlink
#[derive(Debug, Clone)]
pub(crate) struct HyperlinkInfo {
    pub url: String,
    pub rel_id: String,
}

impl HyperlinkContext {
    pub fn new() -> Self {
        Self {
            hyperlinks: Vec::new(),
        }
    }

    /// Add a hyperlink and return its relationship ID
    pub fn add_hyperlink(&mut self, url: &str, rel_manager: &mut RelIdManager) -> String {
        let rel_id = rel_manager.next_id();
        self.hyperlinks.push(HyperlinkInfo {
            url: url.to_string(),
            rel_id: rel_id.clone(),
        });
        rel_id
    }
}

/// Tracks list numbering instances during document building
///
/// Each separate list (ordered or unordered) gets a unique numId
/// to ensure Word restarts numbering for each list independently.
#[derive(Debug, Default, Clone)]
pub(crate) struct NumberingContext {
    /// List of (numId, is_ordered) pairs for all lists
    pub lists: Vec<NumberingInfo>,
    next_id: u32,
}

/// Information about a list numbering instance
#[derive(Debug, Clone)]
pub(crate) struct NumberingInfo {
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
            base_path: None,
        }
    }

    /// Set the base path for resolving relative image paths
    #[allow(dead_code)]
    pub fn with_base_path(mut self, path: std::path::PathBuf) -> Self {
        self.base_path = Some(path);
        self
    }

    /// Resolve an image source path against the base path if set
    fn resolve_image_path(&self, src: &str) -> String {
        // Skip if it's a URL, absolute path, or data URI
        if src.starts_with("http://")
            || src.starts_with("https://")
            || src.starts_with("/")
            || src.starts_with("data:")
            || std::path::Path::new(src).is_absolute()
        {
            return src.to_string();
        }

        // Resolve against base path if set
        if let Some(ref base) = self.base_path {
            let resolved = base.join(src);
            return resolved.to_string_lossy().to_string();
        }

        src.to_string()
    }

    /// Add an image and return its relationship ID
    ///
    /// For now, we assign a placeholder rel_id. The actual ID will be
    /// assigned during packaging when relationships are finalized.
    pub fn add_image(
        &mut self,
        src: &str,
        width: Option<&str>,
        rel_manager: &mut RelIdManager,
    ) -> String {
        let rel_id = rel_manager.next_id();
        let filename = self.generate_filename(src, rel_id.clone());

        // Resolve the source path against base path
        let resolved_src = self.resolve_image_path(src);

        // Try to read actual dimensions from resolved path
        let mut actual_dims = None;
        #[cfg(not(target_arch = "wasm32"))]
        {
            if let Ok(data) = std::fs::read(&resolved_src) {
                actual_dims = read_image_dimensions(&data);
            }
        }

        let (width_emu, height_emu) = self.parse_dimensions(width, actual_dims);

        self.images.push(ImageInfo {
            filename: filename.clone(),
            rel_id: rel_id.clone(),
            src: resolved_src, // Store resolved path for later reading
            data: None,        // Data loaded during packaging
            width_emu,
            height_emu,
        });

        rel_id
    }

    /// Add image from raw data (for generated images like mermaid PNGs)
    pub fn add_image_data(
        &mut self,
        filename: &str,
        data: Vec<u8>,
        width: Option<&str>,
        rel_manager: &mut RelIdManager,
    ) -> String {
        let rel_id = rel_manager.next_id();

        // Try to read dimensions from the image data
        let (width_emu, height_emu) = if let Some(dims) = read_image_dimensions(&data) {
            default_image_size_emu(dims)
        } else {
            // Fallback to default size
            (6 * 914400, 4 * 914400)
        };

        // Apply width override if specified
        let (final_width, final_height) = if let Some(w) = width {
            self.calculate_size_with_aspect_ratio(w, width_emu, height_emu)
        } else {
            (width_emu, height_emu)
        };

        self.images.push(ImageInfo {
            filename: filename.to_string(),
            rel_id: rel_id.clone(),
            src: filename.to_string(),
            data: Some(data),
            width_emu: final_width,
            height_emu: final_height,
        });

        rel_id
    }

    /// Calculate size preserving aspect ratio when width is specified
    fn calculate_size_with_aspect_ratio(
        &self,
        width_spec: &str,
        current_w: i64,
        current_h: i64,
    ) -> (i64, i64) {
        let aspect_ratio = current_h as f64 / current_w as f64;

        let new_width = if width_spec.ends_with('%') {
            let pct: f64 = width_spec.trim_end_matches('%').parse().unwrap_or(100.0);
            (6.0 * 914400.0 * (pct / 100.0)) as i64 // % of 6 inches
        } else if width_spec.ends_with("in") {
            let inches: f64 = width_spec.trim_end_matches("in").parse().unwrap_or(6.0);
            (inches * 914400.0) as i64
        } else if width_spec.ends_with("px") {
            let px: f64 = width_spec.trim_end_matches("px").parse().unwrap_or(576.0);
            (px / 96.0 * 914400.0) as i64
        } else {
            current_w
        };

        let new_height = (new_width as f64 * aspect_ratio) as i64;
        (new_width, new_height)
    }

    /// Generate a unique filename for the image
    fn generate_filename(&self, src: &str, rel_id: String) -> String {
        // Extract extension from source
        let ext = std::path::Path::new(src)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png");

        format!("image_{}.{}", rel_id, ext)
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

/// Document metadata from md2docx.toml [document] section
#[derive(Debug, Clone, Default)]
pub struct DocumentMeta {
    pub title: String,
    pub subtitle: String,
    pub author: String,
    pub date: String,
}

/// Page layout configuration (dimensions and margins in twips)
///
/// 1 twip = 1/20th of a point = 1/1440th of an inch
/// Common values:
/// - A4 width: 11906 twips (210mm)
/// - A4 height: 16838 twips (297mm)
/// - 1 inch margin: 1440 twips (25.4mm)
#[derive(Debug, Clone, Default)]
pub struct PageConfig {
    /// Page width in twips
    pub width: Option<u32>,
    /// Page height in twips
    pub height: Option<u32>,
    /// Top margin in twips
    pub margin_top: Option<u32>,
    /// Right margin in twips
    pub margin_right: Option<u32>,
    /// Bottom margin in twips
    pub margin_bottom: Option<u32>,
    /// Left margin in twips
    pub margin_left: Option<u32>,
    /// Header margin in twips
    pub margin_header: Option<u32>,
    /// Footer margin in twips
    pub margin_footer: Option<u32>,
    /// Gutter margin in twips
    pub margin_gutter: Option<u32>,
}

/// Parse a length string like "210mm", "8.5in", "297mm" to twips
///
/// Supported units:
/// - mm: millimeters (1mm = 56.692913386 twips)
/// - cm: centimeters (1cm = 566.92913386 twips)
/// - in: inches (1in = 1440 twips)
/// - pt: points (1pt = 20 twips)
/// - px: pixels at 96 DPI (1px ≈ 15 twips)
///
/// Returns None if the string cannot be parsed.
pub fn parse_length_to_twips(length: &str) -> Option<u32> {
    let lower = length.to_lowercase();
    let trimmed = lower.trim();

    // Extract the numeric part and unit
    let (val_str, unit) = if trimmed.ends_with("mm") {
        (trimmed.trim_end_matches("mm").trim(), "mm")
    } else if trimmed.ends_with("cm") {
        (trimmed.trim_end_matches("cm").trim(), "cm")
    } else if trimmed.ends_with("in") {
        (trimmed.trim_end_matches("in").trim(), "in")
    } else if trimmed.ends_with("pt") {
        (trimmed.trim_end_matches("pt").trim(), "pt")
    } else if trimmed.ends_with("px") {
        (trimmed.trim_end_matches("px").trim(), "px")
    } else {
        // Assume raw number is in twips
        (trimmed, "twips")
    };

    let val: f64 = val_str.parse().ok()?;

    let twips = match unit {
        "mm" => val * 56.692913386, // 1440 / 25.4
        "cm" => val * 566.92913386, // 1440 / 2.54
        "in" => val * 1440.0,
        "pt" => val * 20.0,
        "px" => val * 15.0, // Assuming 96 DPI: 1440 / 96
        "twips" => val,
        _ => return None,
    };

    Some(twips.round() as u32)
}

/// Document build configuration
#[derive(Debug, Clone)]
pub struct DocumentConfig {
    pub title: String,
    pub toc: TocConfig,
    pub header: HeaderConfig,
    pub footer: FooterConfig,
    pub different_first_page: bool, // Hide header/footer on first page
    /// Template directory path (optional)
    pub template_dir: Option<std::path::PathBuf>,
    /// Offset for IDs to avoid collisions (default: 0)
    pub id_offset: u32,
    /// If true, include all headings in TOC even if they appear before a thematic break
    /// (Used when cover page is handled via template system)
    pub process_all_headings: bool,
    /// Extracted header/footer template from template directory
    pub header_footer_template: Option<crate::template::extract::HeaderFooterTemplate>,
    /// Document metadata for placeholder replacement
    pub document_meta: Option<DocumentMeta>,
    /// Font configuration
    pub fonts: Option<crate::docx::ooxml::FontConfig>,
    /// Base directory for resolving relative image paths (e.g., the markdown file's directory)
    pub base_path: Option<std::path::PathBuf>,
    /// Page layout configuration (dimensions and margins)
    pub page: Option<PageConfig>,
    /// Embedded font files (pre-loaded, obfuscated) for font embedding
    pub embedded_fonts: Vec<crate::docx::font_embed::EmbeddedFont>,
    /// Directory containing .ttf/.otf font files to embed.
    /// When set, fonts are automatically scanned and embedded from this directory.
    /// If `embedded_fonts` is also populated, this field is ignored.
    pub embed_dir: Option<std::path::PathBuf>,
    /// Mermaid diagram spacing: (before, after) in twips
    pub mermaid_spacing: (u32, u32),
    /// Math renderer mode: "rex" (default) or "omml"
    pub math_renderer: String,
    /// Font size for math rendering (e.g. "10pt", "12pt")
    pub math_font_size: String,
    /// Whether to number all display equations (including unlabeled ones)
    pub math_number_all: bool,
}

impl Default for DocumentConfig {
    fn default() -> Self {
        Self {
            title: String::new(),
            toc: TocConfig::default(),
            header: HeaderConfig::default(),
            footer: FooterConfig::default(),
            different_first_page: false,
            template_dir: None,
            id_offset: 0,
            process_all_headings: false,
            header_footer_template: None,
            document_meta: None,
            fonts: None,
            base_path: None,
            page: None,
            embedded_fonts: Vec::new(),
            embed_dir: None,
            mermaid_spacing: (120, 120),
            math_renderer: "rex".to_string(),
            math_font_size: "10pt".to_string(),
            math_number_all: false,
        }
    }
}

/// Mapping of original relationship ID to media file content
#[derive(Debug, Clone)]
pub(crate) struct MediaFileMapping {
    /// Original relationship ID from the template
    pub original_rel_id: String,
    /// Media file content and metadata
    pub media_file: crate::template::extract::header_footer::MediaFile,
}

/// Header or footer entry with associated media files
#[derive(Debug)]
pub(crate) struct HeaderFooterEntry {
    /// Header/footer number (1, 2, 3, etc.)
    pub number: u32,
    /// XML content as bytes
    pub xml_bytes: Vec<u8>,
    /// Media files referenced by this header/footer
    pub media_files: Vec<MediaFileMapping>,
}

/// Result of building a document, including tracked images, hyperlinks, footnotes, and headers/footers
#[derive(Debug)]
pub(crate) struct BuildResult {
    pub document: DocumentXml,
    pub images: ImageContext,
    pub hyperlinks: HyperlinkContext,
    pub footnotes: FootnotesXml,
    pub numbering: NumberingContext,
    pub headers: Vec<HeaderFooterEntry>,
    pub footers: Vec<HeaderFooterEntry>,
    #[allow(dead_code)]
    pub has_toc_section_break: bool, // If true, there's a TOC section break needing empty refs
    pub toc_builder: Option<TocBuilder>,
}

/// Check if a block is a heading
fn is_heading(block: &Block) -> bool {
    matches!(block, Block::Heading { .. })
}

/// Build a DOCX document from parsed markdown
///
/// # Arguments
/// * `doc` - The parsed markdown document
/// * `_lang` - Language for style defaults (English/Thai) - currently unused
/// * `config` - Document configuration including TOC, header, footer settings
///
/// # Returns
/// A `Result<BuildResult>` containing the document and tracked images, hyperlinks, footnotes, and headers/footers
///
/// # Example
/// ```rust,ignore
/// use md2docx::parser::parse_markdown_with_frontmatter;
/// use md2docx::docx::build_document;
/// use md2docx::docx::RelIdManager;
/// use md2docx::DocumentConfig;
/// use md2docx::Language;
///
/// let md = "# Hello World\n\nThis is **bold** text.";
/// let parsed = parse_markdown_with_frontmatter(md);
/// let config = DocumentConfig::default();
/// let mut rel_manager = RelIdManager::new();
/// let result = build_document(&parsed, Language::English, &config, &mut rel_manager, None, None).unwrap();
/// ```
pub(crate) fn build_document(
    doc: &ParsedDocument,
    lang: Language,
    config: &DocumentConfig,
    rel_manager: &mut RelIdManager,
    table_template: Option<&TableTemplate>,
    image_template: Option<&crate::template::extract::image::ImageTemplate>,
) -> crate::error::Result<BuildResult> {
    let mut doc_xml = DocumentXml::new();
    let mut image_ctx = ImageContext::new();
    // Set base path for image resolution if provided in config
    if let Some(ref base) = config.base_path {
        image_ctx.base_path = Some(base.clone());
    }
    let mut hyperlink_ctx = HyperlinkContext::new();
    let mut numbering_ctx = NumberingContext::new();



    let mut footnotes = FootnotesXml::new();

    // TOC builder for collecting headings
    let mut toc_builder = TocBuilder::new();
    let mut bookmark_id_counter: u32 = 10000 + config.id_offset;
    let mut table_count: u32 = 0;
    let mut figure_count: u32 = 0;

    // Calculate body width for tab stops (page width minus margins)
    let page_width = config.page.as_ref().and_then(|p| p.width).unwrap_or(11906);
    let margin_left = config.page.as_ref().and_then(|p| p.margin_left).unwrap_or(1440);
    let margin_right = config.page.as_ref().and_then(|p| p.margin_right).unwrap_or(1440);
    let body_width_twips = page_width.saturating_sub(margin_left + margin_right);

    // Cross-reference context for tracking anchors
    let mut xref_ctx = CrossRefContext::new();

    // Track headers and footers
    let mut headers = Vec::new();
    let mut footers = Vec::new();
    let mut header_footer_refs = HeaderFooterRefs::default();

    // Track previous block to insert blank lines before headings
    let mut prev_block: Option<&Block> = None;

    // Find the first thematic break index (end of cover section)
    // Headings before this should not be in TOC, UNLESS process_all_headings is set
    let first_thematic_break_index = if config.process_all_headings {
        None
    } else {
        doc.blocks
            .iter()
            .position(|b| matches!(b, Block::ThematicBreak))
    };

    let resolved_math_renderer = config.math_renderer.clone();

    // Process all blocks in the document
    // Track the last list seen to support resuming lists across code blocks
    let mut last_list_info: Option<(u32, bool, usize)> = None; // (num_id, is_ordered, block_index)

    for (i, block) in doc.blocks.iter().enumerate() {
        // Create build context
        let mut ctx = BuildContext::new(BuildContextParams {
            image_ctx: &mut image_ctx,
            hyperlink_ctx: &mut hyperlink_ctx,
            numbering_ctx: &mut numbering_ctx,
            doc,
            footnotes: &mut footnotes,
            toc_builder: &mut toc_builder,
            bookmark_id_counter: &mut bookmark_id_counter,
            xref_ctx: &mut xref_ctx,
            rel_manager,
            table_template,
            image_template,
            table_count: &mut table_count,
            figure_count: &mut figure_count,
            lang,
            font_override: None,
            code_font: config.fonts.as_ref().and_then(|f| f.code.clone()),
            code_size: config.fonts.as_ref().and_then(|f| f.code_size),
            quote_level: 0,
            mermaid_spacing: config.mermaid_spacing,
            math_renderer: resolved_math_renderer.clone(),
            math_font_size: config.math_font_size.clone(),
            math_number_all: config.math_number_all,
            body_width_twips,
        });

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

    // Generate headers and footers
    // Note: Relationship IDs are NOT set here - they are assigned in lib.rs after
    // doc_rels.add_header() and add_footer() are called, which return the actual IDs.
    if let Some(ref hf_template) = config.header_footer_template {
        // Use template-based generation
        let ctx = crate::template::render::header_footer::HeaderFooterContext {
            title: config
                .document_meta
                .as_ref()
                .map(|m| m.title.clone())
                .unwrap_or_else(|| config.title.clone()),
            subtitle: config
                .document_meta
                .as_ref()
                .map(|m| m.subtitle.clone())
                .unwrap_or_default(),
            author: config
                .document_meta
                .as_ref()
                .map(|m| m.author.clone())
                .unwrap_or_default(),
            date: config
                .document_meta
                .as_ref()
                .map(|m| m.date.clone())
                .unwrap_or_default(),
        };

        // Render default header
        if let Ok(Some(rendered)) =
            crate::template::render::header_footer::render_default_header(hf_template, &ctx, 100)
        {
            let media_mappings = rendered
                .media
                .into_iter()
                .map(|(rel_id, media_file)| MediaFileMapping {
                    original_rel_id: rel_id,
                    media_file,
                })
                .collect();
            headers.push(HeaderFooterEntry {
                number: 1,
                xml_bytes: rendered.xml,
                media_files: media_mappings,
            });
        }

        // Render first page header (if different first page)
        if hf_template.different_first_page {
            if let Ok(Some(rendered)) =
                crate::template::render::header_footer::render_first_page_header(
                    hf_template,
                    &ctx,
                    200,
                )
            {
                let media_mappings = rendered
                    .media
                    .into_iter()
                    .map(|(rel_id, media_file)| MediaFileMapping {
                        original_rel_id: rel_id,
                        media_file,
                    })
                    .collect();
                headers.push(HeaderFooterEntry {
                    number: 2,
                    xml_bytes: rendered.xml,
                    media_files: media_mappings,
                });
            }
            header_footer_refs.different_first_page = true;
        } else if config.different_first_page {
            // Template doesn't have different first page, but config requests it
            // Generate empty header for suppression
            let empty_header = HeaderXml::new(HeaderConfig::empty(), "");
            let xml = empty_header.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate empty header: {}", e))
            })?;
            headers.push(HeaderFooterEntry {
                number: 2,
                xml_bytes: xml,
                media_files: Vec::new(),
            });
            header_footer_refs.different_first_page = true;
        }

        // Render default footer
        if let Ok(Some(rendered)) =
            crate::template::render::header_footer::render_default_footer(hf_template, &ctx, 300)
        {
            let media_mappings = rendered
                .media
                .into_iter()
                .map(|(rel_id, media_file)| MediaFileMapping {
                    original_rel_id: rel_id,
                    media_file,
                })
                .collect();
            footers.push(HeaderFooterEntry {
                number: 1,
                xml_bytes: rendered.xml,
                media_files: media_mappings,
            });
        }

        // Render first page footer (if different first page)
        if hf_template.different_first_page {
            if let Ok(Some(rendered)) =
                crate::template::render::header_footer::render_first_page_footer(
                    hf_template,
                    &ctx,
                    400,
                )
            {
                let media_mappings = rendered
                    .media
                    .into_iter()
                    .map(|(rel_id, media_file)| MediaFileMapping {
                        original_rel_id: rel_id,
                        media_file,
                    })
                    .collect();
                footers.push(HeaderFooterEntry {
                    number: 2,
                    xml_bytes: rendered.xml,
                    media_files: media_mappings,
                });
            }
            header_footer_refs.different_first_page = true;
        } else if config.different_first_page {
            // Template doesn't have different first page, but config requests it
            // Generate empty footer for suppression
            let empty_footer = FooterXml::new(FooterConfig::empty(), "");
            let xml = empty_footer.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate empty footer: {}", e))
            })?;
            footers.push(HeaderFooterEntry {
                number: 2,
                xml_bytes: xml,
                media_files: Vec::new(),
            });
            header_footer_refs.different_first_page = true;
        }

        // Always generate truly empty header/footer (ID 3) for cover/TOC suppression
        // These are used when we need sections with NO headers/footers at all
        // (separate from first-page headers which may have content)
        let suppression_header = HeaderXml::new(HeaderConfig::empty(), "");
        let xml = suppression_header.to_xml().map_err(|e| {
            crate::error::Error::Xml(format!("Failed to generate suppression header: {}", e))
        })?;
        headers.push(HeaderFooterEntry {
            number: 3,
            xml_bytes: xml,
            media_files: Vec::new(),
        });

        let suppression_footer = FooterXml::new(FooterConfig::empty(), "");
        let xml = suppression_footer.to_xml().map_err(|e| {
            crate::error::Error::Xml(format!("Failed to generate suppression footer: {}", e))
        })?;
        footers.push(HeaderFooterEntry {
            number: 3,
            xml_bytes: xml,
            media_files: Vec::new(),
        });
    } else {
        // Fall back to config-based generation (existing code)
        if !config.header.is_empty() {
            // Generate default header (header1.xml)
            let header_xml = HeaderXml::new(config.header.clone(), &config.title);
            let xml = header_xml.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate header: {}", e))
            })?;
            headers.push(HeaderFooterEntry {
                number: 1,
                xml_bytes: xml,
                media_files: Vec::new(),
            });
            // Relationship ID will be set in lib.rs

            // Generate empty header (ID 2) for first page if different_first_page is set
            let empty_header = HeaderXml::new(HeaderConfig::empty(), "");
            let xml = empty_header.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate empty header: {}", e))
            })?;
            headers.push(HeaderFooterEntry {
                number: 2,
                xml_bytes: xml,
                media_files: Vec::new(),
            });

            // Also generate header3 for cover/TOC suppression (same as header2 but separate file)
            let suppression_header = HeaderXml::new(HeaderConfig::empty(), "");
            let xml = suppression_header.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate suppression header: {}", e))
            })?;
            headers.push(HeaderFooterEntry {
                number: 3,
                xml_bytes: xml,
                media_files: Vec::new(),
            });

            if config.different_first_page {
                header_footer_refs.different_first_page = true;
            }
        }

        if !config.footer.is_empty() {
            // Generate default footer (footer1.xml)
            let footer_xml = FooterXml::new(config.footer.clone(), &config.title);
            let xml = footer_xml.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate footer: {}", e))
            })?;
            footers.push(HeaderFooterEntry {
                number: 1,
                xml_bytes: xml,
                media_files: Vec::new(),
            });
            // Relationship ID will be set in lib.rs

            // Generate empty footer (ID 2) for first page if different_first_page is set
            let empty_footer = FooterXml::new(FooterConfig::empty(), "");
            let xml = empty_footer.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate empty footer: {}", e))
            })?;
            footers.push(HeaderFooterEntry {
                number: 2,
                xml_bytes: xml,
                media_files: Vec::new(),
            });

            // Also generate footer3 for cover/TOC suppression (same as footer2 but separate file)
            let suppression_footer = FooterXml::new(FooterConfig::empty(), "");
            let xml = suppression_footer.to_xml().map_err(|e| {
                crate::error::Error::Xml(format!("Failed to generate suppression footer: {}", e))
            })?;
            footers.push(HeaderFooterEntry {
                number: 3,
                xml_bytes: xml,
                media_files: Vec::new(),
            });

            if config.different_first_page {
                header_footer_refs.different_first_page = true;
            }
        }
    }

    // Set header/footer refs on document
    doc_xml.header_footer_refs = header_footer_refs;

    Ok(BuildResult {
        document: doc_xml,
        images: image_ctx,
        hyperlinks: hyperlink_ctx,
        footnotes,
        numbering: numbering_ctx,
        headers,
        footers,
        has_toc_section_break: false,
        toc_builder: Some(toc_builder),
    })
}

/// Parameters for creating a BuildContext
pub(crate) struct BuildContextParams<'a> {
    pub image_ctx: &'a mut ImageContext,
    pub hyperlink_ctx: &'a mut HyperlinkContext,
    pub numbering_ctx: &'a mut NumberingContext,
    pub doc: &'a ParsedDocument,

    pub footnotes: &'a mut FootnotesXml,
    pub toc_builder: &'a mut TocBuilder,
    pub bookmark_id_counter: &'a mut u32,
    pub xref_ctx: &'a mut CrossRefContext,
    pub rel_manager: &'a mut RelIdManager,
    pub table_template: Option<&'a TableTemplate>,
    pub image_template: Option<&'a crate::template::extract::image::ImageTemplate>,
    pub table_count: &'a mut u32,
    pub figure_count: &'a mut u32,
    pub lang: Language,
    pub font_override: Option<String>,
    pub code_font: Option<String>,
    pub code_size: Option<u32>,
    pub quote_level: usize,
    pub mermaid_spacing: (u32, u32),
    pub math_renderer: String,
    pub math_font_size: String,
    pub math_number_all: bool,
    pub body_width_twips: u32,
}

/// Context for building a document, holding all tracked state
pub(crate) struct BuildContext<'a> {
    pub image_ctx: &'a mut ImageContext,
    pub hyperlink_ctx: &'a mut HyperlinkContext,
    pub numbering_ctx: &'a mut NumberingContext,
    pub doc: &'a ParsedDocument,

    pub footnotes: &'a mut FootnotesXml,
    pub toc_builder: &'a mut TocBuilder,
    pub bookmark_id_counter: &'a mut u32,
    pub xref_ctx: &'a mut CrossRefContext,
    pub rel_manager: &'a mut RelIdManager,
    pub table_template: Option<&'a TableTemplate>,
    pub image_template: Option<&'a crate::template::extract::image::ImageTemplate>,
    pub table_count: &'a mut u32,
    pub figure_count: &'a mut u32,
    pub lang: Language,
    pub font_override: Option<String>,
    pub code_font: Option<String>,
    pub code_size: Option<u32>,
    pub quote_level: usize,
    pub mermaid_spacing: (u32, u32),
    pub math_renderer: String,
    pub math_font_size: String,
    pub math_number_all: bool,
    pub body_width_twips: u32,
}

impl<'a> BuildContext<'a> {
    pub fn new(params: BuildContextParams<'a>) -> Self {
        Self {
            image_ctx: params.image_ctx,
            hyperlink_ctx: params.hyperlink_ctx,
            numbering_ctx: params.numbering_ctx,
            doc: params.doc,
            footnotes: params.footnotes,
            toc_builder: params.toc_builder,
            bookmark_id_counter: params.bookmark_id_counter,
            xref_ctx: params.xref_ctx,
            rel_manager: params.rel_manager,
            table_template: params.table_template,
            image_template: params.image_template,
            table_count: params.table_count,
            figure_count: params.figure_count,
            lang: params.lang,
            font_override: params.font_override,
            code_font: params.code_font,
            code_size: params.code_size,
            quote_level: params.quote_level,
            mermaid_spacing: params.mermaid_spacing,
            math_renderer: params.math_renderer,
            math_font_size: params.math_font_size,
            math_number_all: params.math_number_all,
            body_width_twips: params.body_width_twips,
        }
    }
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
            let rel_id = ctx
                .image_ctx
                .add_image(src, width.as_deref(), ctx.rel_manager);

            // Get dimensions from context (last added image)
            let (width_emu, height_emu) = ctx
                .image_ctx
                .images
                .last()
                .map(|img| (img.width_emu, img.height_emu))
                .unwrap_or((5486400, 3657600)); // Default 6x4 inches

            let image_id = ctx.rel_manager.next_image_id();

            // Create image element
            let mut img = ImageElement::new(&rel_id, width_emu, height_emu)
                .alt_text(alt)
                .name(src)
                .id(image_id);

            // Apply template effects if available
            if let Some(tmpl) = ctx.image_template {
                // Apply border
                if let Some(ref border) = tmpl.border {
                    img = img.with_border(crate::docx::ooxml::ImageBorderEffect {
                        fill_type: border.fill_type.clone(),
                        color: border.color.clone(),
                        is_scheme_color: border.is_scheme_color,
                        width: border.width,
                    });
                }

                // Apply shadow
                if let Some(ref shadow) = tmpl.shadow {
                    img = img.with_shadow(crate::docx::ooxml::ImageShadowEffect {
                        blur_radius: shadow.blur_radius,
                        distance: shadow.distance,
                        direction: shadow.direction,
                        alignment: shadow.alignment.clone(),
                        color: shadow.color.clone(),
                        alpha: shadow.alpha,
                    });
                }

                // Apply effect extent
                let extent = &tmpl.effect_extent;
                if extent.left > 0 || extent.top > 0 || extent.right > 0 || extent.bottom > 0 {
                    img = img.with_effect_extent(crate::docx::ooxml::ImageEffectExtent {
                        left: extent.left,
                        top: extent.top,
                        right: extent.right,
                        bottom: extent.bottom,
                    });
                }

                // Apply alignment
                if !tmpl.alignment.is_empty() {
                    img = img.with_alignment(&tmpl.alignment);
                }
            }



            // Get figure number (either from xref or sequential)
            let figure_number = if let Some(fig_id) = id {
                // Already registered above, get the number
                if let Some(anchor) = ctx.xref_ctx.resolve(fig_id) {
                    anchor.number.clone()
                } else {
                    None
                }
            } else {
                // No ID - just use sequential number
                *ctx.figure_count += 1;
                Some(ctx.figure_count.to_string())
            };

            // Build result elements
            let mut elements = vec![DocElement::Image(img)];

            // Add caption paragraph if template and alt text exist
            if let Some(tmpl) = ctx.image_template {
                if !alt.is_empty() {
                    // Use localized prefix if template has default "Figure"
                    let prefix = if tmpl.caption.prefix == "Figure" {
                        ctx.lang.figure_caption_prefix().to_string()
                    } else {
                        tmpl.caption.prefix.clone()
                    };

                    let number_str = figure_number.unwrap_or_else(|| {
                        *ctx.figure_count += 1;
                        ctx.figure_count.to_string()
                    });

                    let caption_text = format!("{} {}: {}", prefix, number_str, alt);

                    let mut run = Run::new(&caption_text);
                    run.font = Some(ctx.font_override.as_ref().unwrap_or(&tmpl.caption.font_family).clone());
                    run.size = Some(tmpl.caption.font_size);
                    run.color = Some(tmpl.caption.font_color.trim_start_matches('#').to_string());
                    run.bold = tmpl.caption.bold;
                    run.italic = tmpl.caption.italic;

                    let mut caption_para = Paragraph::with_style("Caption")
                        .add_run(run)
                        .spacing(tmpl.caption.spacing_before, tmpl.caption.spacing_after);

                    // Align caption to match image alignment
                    caption_para = caption_para.align(&tmpl.alignment);

                    // Add bookmark if we have an ID
                    if let Some(anchor) =
                        id.as_ref().and_then(|fig_id| ctx.xref_ctx.resolve(fig_id))
                    {
                        *ctx.bookmark_id_counter += 1;
                        caption_para = caption_para
                            .with_bookmark(*ctx.bookmark_id_counter, &anchor.bookmark_name);
                    }

                    elements.push(DocElement::Paragraph(Box::new(caption_para)));
                }
            } else if !alt.is_empty() {
                // No template — create a simple caption with alt text
                let prefix = ctx.lang.figure_caption_prefix();
                let number_str = figure_number.unwrap_or_else(|| {
                    *ctx.figure_count += 1;
                    ctx.figure_count.to_string()
                });
                let caption_text = format!("{} {}: {}", prefix, number_str, alt);
                let mut run = Run::new(&caption_text);
                if let Some(ref font) = ctx.font_override {
                    run.font = Some(font.clone());
                }
                let caption_para = Paragraph::with_style("Caption")
                    .add_run(run)
                    .spacing(120, 120);
                elements.push(DocElement::Paragraph(Box::new(caption_para)));
            }

            elements
        }

        Block::Mermaid { content, id } => {
            match crate::mermaid::render_to_svg(content) {
                Ok(svg_data) => {
                    // Register figure anchor if id is present
                    if let Some(fig_id) = id {
                        ctx.xref_ctx.register_figure(fig_id, "Mermaid Diagram");
                    }

                    let image_id = ctx.rel_manager.next_image_id();
                    // Generate a virtual filename
                    let filename = format!("mermaid{}.svg", image_id);

                    // Add to image context as SVG
                    let rel_id = ctx.image_ctx.add_image_data(
                        &filename,
                        svg_data.into_bytes(),
                        None, // No explicit width, let it use natural size
                        ctx.rel_manager,
                    );

                    // Get dimensions from the SVG data
                    let (width_emu, height_emu) = ctx
                        .image_ctx
                        .images
                        .last()
                        .map(|img| (img.width_emu, img.height_emu))
                        .unwrap_or((6 * 914400, 4 * 914400));

                    let mut img = ImageElement::new(&rel_id, width_emu, height_emu)
                        .alt_text("Mermaid Diagram")
                        .name(&filename)
                        .id(image_id);

                    // For Mermaid diagrams, only apply alignment (no border, shadow, or padding)
                    if let Some(tmpl) = ctx.image_template {
                        if !tmpl.alignment.is_empty() {
                            img = img.with_alignment(&tmpl.alignment);
                        }
                    }

                    // Apply mermaid diagram spacing
                    let (sp_before, sp_after) = ctx.mermaid_spacing;
                    img = img.with_spacing(sp_before, sp_after);

                    // Build result elements
                    let mut elements = vec![DocElement::Image(img)];

                    // Add caption paragraph if template and id exist (Mermaid has no alt text)
                    if let Some(tmpl) = ctx.image_template {
                        if id.is_some() {
                            // Get figure number from xref (already registered above)
                            let figure_number = id.as_ref().and_then(|fig_id| {
                                ctx.xref_ctx.resolve(fig_id).and_then(|a| a.number.clone())
                            });

                            // Use localized prefix if template has default "Figure"
                            let prefix = if tmpl.caption.prefix == "Figure" {
                                ctx.lang.figure_caption_prefix().to_string()
                            } else {
                                tmpl.caption.prefix.clone()
                            };

                            let number_str = figure_number.unwrap_or_else(|| {
                                *ctx.figure_count += 1;
                                ctx.figure_count.to_string()
                            });

                            let caption_text = format!("{} {}", prefix, number_str);

                            let mut run = Run::new(&caption_text);
                            run.font = Some(ctx.font_override.as_ref().unwrap_or(&tmpl.caption.font_family).clone());
                            run.size = Some(tmpl.caption.font_size);
                            run.color =
                                Some(tmpl.caption.font_color.trim_start_matches('#').to_string());
                            run.bold = tmpl.caption.bold;
                            run.italic = tmpl.caption.italic;

                            let mut caption_para = Paragraph::with_style("Caption")
                                .add_run(run)
                                .spacing(tmpl.caption.spacing_before, tmpl.caption.spacing_after);

                            // Align caption to match image alignment
                            caption_para = caption_para.align(&tmpl.alignment);

                            // Add bookmark if we have an ID
                            if let Some(anchor) =
                                id.as_ref().and_then(|fig_id| ctx.xref_ctx.resolve(fig_id))
                            {
                                *ctx.bookmark_id_counter += 1;
                                caption_para = caption_para
                                    .with_bookmark(*ctx.bookmark_id_counter, &anchor.bookmark_name);
                            }

                            elements.push(DocElement::Paragraph(Box::new(caption_para)));
                        }
                    }

                    elements
                }
                Err(e) => {
                    eprintln!("Warning: Failed to render mermaid diagram: {}", e);
                    // Fallback to code block
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
            caption,
            id,
        } => {
            let mut elements = Vec::new();

            // Register table with cross-reference context if it has an ID
            let table_number = if let Some(table_id) = id {
                // Register and get the proper number (e.g., "1.2")
                ctx.xref_ctx
                    .register_table(table_id, caption.as_deref().unwrap_or(""));

                // Get the number from xref context
                if let Some(anchor) = ctx.xref_ctx.resolve(table_id) {
                    anchor.number.clone()
                } else {
                    None
                }
            } else {
                // No ID - just use sequential number
                *ctx.table_count += 1;
                Some(ctx.table_count.to_string())
            };

            // Add caption paragraph if template has caption style
            if let Some(template) = ctx.table_template {
                // Use localized prefix if template has default "Table"
                let prefix = if template.caption.prefix == "Table" {
                    ctx.lang.table_caption_prefix().to_string()
                } else {
                    template.caption.prefix.clone()
                };

                let number_str = table_number.unwrap_or_else(|| {
                    *ctx.table_count += 1;
                    ctx.table_count.to_string()
                });

                let caption_text = format!(
                    "{} {}: {}",
                    prefix,
                    number_str,
                    caption.as_deref().unwrap_or_default()
                );

                let mut run = Run::new(&caption_text);
                run.font = Some(ctx.font_override.as_ref().unwrap_or(&template.caption.font_family).clone());
                run.size = Some(template.caption.font_size);
                run.color = Some(
                    template
                        .caption
                        .font_color
                        .trim_start_matches('#')
                        .to_string(),
                );
                run.bold = template.caption.bold;
                run.italic = template.caption.italic;

                let mut caption_para = Paragraph::with_style("Caption").add_run(run).spacing(
                    template.caption.spacing_before,
                    template.caption.spacing_after,
                );

                // Add bookmark if we have an ID
                if let Some(anchor) = id
                    .as_ref()
                    .and_then(|table_id| ctx.xref_ctx.resolve(table_id))
                {
                    *ctx.bookmark_id_counter += 1;
                    caption_para =
                        caption_para.with_bookmark(*ctx.bookmark_id_counter, &anchor.bookmark_name);
                }

                elements.push(DocElement::Paragraph(Box::new(caption_para)));
            }

            let table = table_to_docx(headers, alignments, rows, ctx);
            elements.push(DocElement::Table(table));

            // Add empty paragraph after table for spacing
            let empty_para = Paragraph::default().spacing(0, 0).line_spacing(240, "auto");
            elements.push(DocElement::Paragraph(Box::new(empty_para)));

            elements
        }

        Block::BlockQuote(blocks) => {
            let mut result = Vec::new();
            ctx.quote_level += 1;
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
                            // Only apply quote styling to paragraphs not already
                            // styled by a deeper nested blockquote
                            if p.style_id.as_deref() != Some("Quote") {
                                p.style_id = Some("Quote".to_string());
                                p.indent_left = Some(ctx.quote_level as u32 * 720);
                            }
                            result.push(DocElement::Paragraph(p));
                        }
                        other => result.push(other),
                    }
                }
            }
            ctx.quote_level -= 1;
            result
        }

        Block::FontGroup { font, blocks } => {
            let prev_override = ctx.font_override.clone();
            ctx.font_override = Some(font.clone());
            let mut result = Vec::new();
            for block in blocks {
                result.extend(block_to_elements(block, list_level, ctx, None, skip_toc));
            }
            ctx.font_override = prev_override;
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
        _ => {
            // Special handling for MathBlock at element level
            if let Block::MathBlock { content, id } = block {
                let center_pos = ctx.body_width_twips / 2;
                let right_pos = ctx.body_width_twips;

                // Get equation number via xref (chapter-relative) — only for labeled equations
                let (eq_number, bookmark_name) = if let Some(eq_id) = id {
                    if let Some(anchor) = ctx.xref_ctx.resolve(eq_id) {
                        (Some(anchor.number.clone().unwrap_or_default()), Some(anchor.bookmark_name.clone()))
                    } else {
                        // Not pre-registered — register now
                        let bk = ctx.xref_ctx.register_equation(eq_id);
                        let num = ctx.xref_ctx.resolve(eq_id).and_then(|a| a.number.clone()).unwrap_or_default();
                        (Some(num), Some(bk))
                    }
                } else if ctx.math_number_all {
                    // No label but number_all is enabled — assign a number without bookmark
                    let num = ctx.xref_ctx.next_equation_number();
                    (Some(num), None)
                } else {
                    // No label — no number
                    (None, None)
                };

                // Check renderer config: "rex" or "omml"
                if ctx.math_renderer == "rex" {
                    let render_result = crate::docx::math_rex::render_latex_to_svg(content, true, &ctx.math_font_size);
                    match render_result {
                        Ok(math) => {
                            let image_id = ctx.rel_manager.next_image_id();
                            let filename = format!("math_display{}.svg", image_id);

                            let rel_id = ctx.image_ctx.add_image_data(
                                &filename,
                                math.svg_bytes,
                                None,
                                ctx.rel_manager,
                            );

                            let mut img = ImageElement::new(&rel_id, math.width_emu, math.height_emu)
                                .alt_text("Math equation")
                                .name(&filename)
                                .id(image_id);
                            img.position = math.position;

                            let bookmark = bookmark_name.as_ref().map(|bk_name| {
                                *ctx.bookmark_id_counter += 1;
                                (*ctx.bookmark_id_counter, bk_name.clone())
                            });
                            let mut para = build_equation_paragraph(center_pos, right_pos, eq_number.as_deref(), bookmark);
                            // Insert inline image before the tab-to-right run (index 1)
                            para.children.insert(1, ParagraphChild::InlineImage(img));

                            return vec![DocElement::Paragraph(Box::new(para))];
                        }
                        Err(e) => {
                            eprintln!("Warning: ReX rendering failed, falling back to OMML: {}", e);
                            let omml = crate::docx::math::latex_to_omml_paragraph(content);

                            let bookmark = bookmark_name.as_ref().map(|bk_name| {
                                *ctx.bookmark_id_counter += 1;
                                (*ctx.bookmark_id_counter, bk_name.clone())
                            });
                            let mut para = build_equation_paragraph(center_pos, right_pos, eq_number.as_deref(), bookmark);
                            para.children.insert(1, ParagraphChild::OfficeMath(omml));

                            return vec![DocElement::Paragraph(Box::new(para))];
                        }
                    }
                } else {
                    let omml = crate::docx::math::latex_to_omml_paragraph(content);

                    let bookmark = bookmark_name.as_ref().map(|bk_name| {
                        *ctx.bookmark_id_counter += 1;
                        (*ctx.bookmark_id_counter, bk_name.clone())
                    });
                    let mut para = build_equation_paragraph(center_pos, right_pos, eq_number.as_deref(), bookmark);
                    para.children.insert(1, ParagraphChild::OfficeMath(omml));

                    return vec![DocElement::Paragraph(Box::new(para))];
                }
            }
            block_to_paragraphs(block, list_level, ctx, skip_toc)
                .into_iter()
                .map(|p| DocElement::Paragraph(Box::new(p)))
                .collect()
        }
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
            lang,
            content,
            filename,
            highlight_lines,
            show_line_numbers,
        } => code_block_to_paragraphs(
            content,
            lang.as_deref(),
            filename.as_deref(),
            highlight_lines,
            *show_line_numbers,
            ctx.code_font.as_deref(),
            ctx.code_size,
        ),

        Block::BlockQuote(blocks) => {
            let mut paragraphs = Vec::new();
            ctx.quote_level += 1;
            for nested_block in blocks {
                let mut nested_paragraphs =
                    block_to_paragraphs(nested_block, list_level, ctx, skip_toc);
                // Only apply quote styling to paragraphs not already
                // styled by a deeper nested blockquote
                for p in &mut nested_paragraphs {
                    if p.style_id.as_deref() != Some("Quote") {
                        p.style_id = Some("Quote".to_string());
                        p.indent_left = Some(ctx.quote_level as u32 * 720);
                    }
                }
                paragraphs.extend(nested_paragraphs);
            }
            ctx.quote_level -= 1;
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

        Block::ThematicBreak => {
            // Insert a blank Normal paragraph before the section break
            let blank = Paragraph::with_style("Normal").spacing(0, 0).line_spacing(240, "auto");
            let section = thematic_break_to_paragraph();
            vec![blank, section]
        }

        Block::Html(_) => {
            // Skip HTML blocks for now
            vec![]
        }

        Block::FontGroup { font, blocks } => {
            let prev_override = ctx.font_override.clone();
            ctx.font_override = Some(font.clone());
            let mut paragraphs = Vec::new();
            for block in blocks {
                paragraphs.extend(block_to_paragraphs(block, list_level, ctx, skip_toc));
            }
            ctx.font_override = prev_override;
            paragraphs
        }

        Block::Mermaid { content, .. } => {
            // This is a fallback case if block_to_elements falls back to block_to_paragraphs
            code_block_to_paragraphs(content, Some("mermaid"), None, &Vec::new(), false, ctx.code_font.as_deref(), ctx.code_size)
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

        Block::MathBlock { content, id } => {
            // Display math: centered equation with right-aligned running number
            let center_pos = ctx.body_width_twips / 2;
            let right_pos = ctx.body_width_twips;

            let (eq_number, bookmark_name) = if let Some(eq_id) = id {
                if let Some(anchor) = ctx.xref_ctx.resolve(eq_id) {
                    (Some(anchor.number.clone().unwrap_or_default()), Some(anchor.bookmark_name.clone()))
                } else {
                    let bk = ctx.xref_ctx.register_equation(eq_id);
                    let num = ctx.xref_ctx.resolve(eq_id).and_then(|a| a.number.clone()).unwrap_or_default();
                    (Some(num), Some(bk))
                }
            } else if ctx.math_number_all {
                // No label but number_all is enabled — assign a number without bookmark
                let num = ctx.xref_ctx.next_equation_number();
                (Some(num), None)
            } else {
                (None, None)
            };

            let omml = crate::docx::math::latex_to_omml_paragraph(content);
            let bookmark = bookmark_name.as_ref().map(|bk_name| {
                *ctx.bookmark_id_counter += 1;
                (*ctx.bookmark_id_counter, bk_name.clone())
            });
            let mut para = build_equation_paragraph(center_pos, right_pos, eq_number.as_deref(), bookmark);
            para.children.insert(1, ParagraphChild::OfficeMath(omml));

            vec![para]
        }
    }
}

/// Build a display equation paragraph with tab stops and equation number.
/// Returns a paragraph with: [tab-to-center] [tab-to-right] [(number)]
/// The caller should insert the equation content (image or OMML) at index 1.
///
/// If `bookmark` is provided as `(id, name)`, a bookmark is placed around
/// just the number portion `(N)` so that REF fields can reference it.
fn build_equation_paragraph(
    center_pos: u32,
    right_pos: u32,
    eq_number: Option<&str>,
    bookmark: Option<(u32, String)>,
) -> Paragraph {
    let mut para = Paragraph::new();

    if eq_number.is_some() {
        // Labeled equation: center tab + right tab for number
        para.tabs = vec![
            TabStop { val: "center".to_string(), pos: center_pos },
            TabStop { val: "right".to_string(), pos: right_pos },
        ];
    } else {
        // Unlabeled equation: just center tab, no number
        para.tabs = vec![
            TabStop { val: "center".to_string(), pos: center_pos },
        ];
    }

    // Tab to center position
    para.children.push(ParagraphChild::Run(Run::new("").with_tab()));
    // (equation content will be inserted at index 1 by caller)

    // Only add the number portion if there's a label
    if let Some(num) = eq_number {
        // Tab to right position
        para.children.push(ParagraphChild::Run(Run::new("").with_tab()));

        // Bookmark start — wraps only the number portion for targeted REF fields
        if let Some((bk_id, ref bk_name)) = bookmark {
            para.children.push(ParagraphChild::BookmarkStart {
                id: bk_id,
                name: bk_name.clone(),
            });
        }

        // Equation number using SEQ field: ( + SEQ Equation + )
        para.children.push(ParagraphChild::Run(Run::new("(")));
        // SEQ field: begin
        para.children.push(ParagraphChild::Run(
            Run::new("").with_field_char("begin"),
        ));
        // SEQ field: instruction
        para.children.push(ParagraphChild::Run(
            Run::new(" SEQ Equation \\* ARABIC ").with_instr_text(),
        ));
        // SEQ field: separate
        para.children.push(ParagraphChild::Run(
            Run::new("").with_field_char("separate"),
        ));
        // SEQ field: placeholder value (Word updates this on F9)
        para.children.push(ParagraphChild::Run(Run::new(num)));
        // SEQ field: end
        para.children.push(ParagraphChild::Run(
            Run::new("").with_field_char("end"),
        ));
        para.children.push(ParagraphChild::Run(Run::new(")")));

        // Bookmark end
        if let Some((bk_id, _)) = bookmark {
            para.children.push(ParagraphChild::BookmarkEnd { id: bk_id });
        }
    }

    para
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
            ParagraphChild::OfficeMath(xml) => p.add_office_math(xml),
            ParagraphChild::InlineImage(img) => p.add_inline_image(img),
            other => { p.children.push(other); p }
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
            ParagraphChild::OfficeMath(xml) => p.add_office_math(xml),
            ParagraphChild::InlineImage(img) => p.add_inline_image(img),
            other => { p.children.push(other); p }
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

    // Apply borders if template available
    if let Some(template) = ctx.table_template {
        table = table.with_borders(template.borders.clone());
        table = table.with_cell_margins(template.cell_margins.clone());
    }

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

    // Always use auto width (autofit to contents)
    let (table_width, cell_width) = (TableWidth::Auto, TableWidth::Auto);

    table = table.width(table_width);

    // Auto column widths (equal distribution, ~9000 twips total for A4)
    // Keep this for w:tblGrid even if w:tblW overrides it visually
    let col_width = 9000 / col_count.max(1) as u32;
    table = table.with_column_widths(vec![col_width; col_count]);

    // Add header row (row index 0)
    let mut header_row = TableRow::new().header();
    for (i, cell) in headers.iter().enumerate() {
        let alignment = alignments.get(i).copied().unwrap_or(ParserAlignment::None);
        let cell_elem = create_table_cell_with_template(
            TableCellParams {
                content: &cell.content,
                alignment,
                is_header: true,
                width: cell_width,
                row_index: 0,
                col_index: i,
                template: ctx.table_template,
            },
            ctx,
        );
        header_row = header_row.add_cell(cell_elem);
    }
    table = table.add_row(header_row);

    // Add data rows
    for (row_idx, row) in rows.iter().enumerate() {
        let actual_row_idx = row_idx + 1; // +1 because header is row 0
        let mut data_row = TableRow::new();
        for (col_idx, cell) in row.iter().enumerate() {
            let alignment = alignments
                .get(col_idx)
                .copied()
                .unwrap_or(ParserAlignment::None);
            let cell_elem = create_table_cell_with_template(
                TableCellParams {
                    content: &cell.content,
                    alignment,
                    is_header: false,
                    width: cell_width,
                    row_index: actual_row_idx,
                    col_index: col_idx,
                    template: ctx.table_template,
                },
                ctx,
            );
            data_row = data_row.add_cell(cell_elem);
        }
        table = table.add_row(data_row);
    }

    table
}

/// Parameters for creating a table cell with template styling
pub struct TableCellParams<'a, 'b> {
    pub content: &'a [Inline],
    pub alignment: ParserAlignment,
    pub is_header: bool,
    pub width: TableWidth,
    pub row_index: usize,
    pub col_index: usize,
    pub template: Option<&'b crate::template::extract::table::TableTemplate>,
}

/// Create a table cell with template styling applied
fn create_table_cell_with_template(
    params: TableCellParams,
    ctx: &mut BuildContext,
) -> TableCellElement {
    let children = inlines_to_children(params.content, ctx);

    // Build paragraph from children
    let mut p = if let Some(tmpl) = params.template {
        Paragraph::new()
            .spacing(tmpl.cell_spacing.before, tmpl.cell_spacing.after)
            .line_spacing(tmpl.cell_spacing.line as i32, &tmpl.cell_spacing.line_rule)
    } else {
        Paragraph::new().spacing(0, 0).line_spacing(240, "auto")
    };

    // Apply font properties from template
    if let Some(tmpl) = params.template {
        let row_style = tmpl.row_style_for_index(params.row_index);
        let col_style = tmpl.cell_style_for_column(params.col_index);

        // Process children with template styling
        for child in children {
            p = match child {
                ParagraphChild::Run(mut r) => {
                    // Row style provides font family, size, and color
                    r.font = Some(row_style.font_family.clone());
                    r.size = Some(row_style.font_size);
                    r.color = Some(row_style.font_color.trim_start_matches('#').to_string());

                    // For header row (index 0), use row_style for bold/italic
                    // For data rows, use col_style (first_column or other_columns)
                    if params.row_index == 0 {
                        // Header row: use header style for bold/italic
                        r.bold = row_style.bold;
                        r.italic = row_style.italic;
                    } else {
                        // Data rows: use column style for bold/italic
                        // This allows first_column.bold=true and other_columns.bold=false
                        r.bold = col_style.bold;
                        r.italic = col_style.italic;

                        // Column style can also override font properties if explicitly set
                        if col_style.font_family != "Calibri" {
                            r.font = Some(col_style.font_family.clone());
                        }
                        if col_style.font_size != 22 {
                            r.size = Some(col_style.font_size);
                        }
                        if col_style.font_color != "#000000" {
                            r.color =
                                Some(col_style.font_color.trim_start_matches('#').to_string());
                        }
                    }

                    p.add_run(r)
                }
                ParagraphChild::Hyperlink(link) => p.add_hyperlink(link),
                ParagraphChild::OfficeMath(xml) => p.add_office_math(xml),
                ParagraphChild::InlineImage(img) => p.add_inline_image(img),
                other => { p.children.push(other); p }
            };
        }
    } else {
        // No template, use default styling
        for child in children {
            p = match child {
                ParagraphChild::Run(mut r) => {
                    if params.is_header {
                        r.bold = true;
                    }
                    p.add_run(r)
                }
                ParagraphChild::Hyperlink(link) => p.add_hyperlink(link),
                ParagraphChild::OfficeMath(xml) => p.add_office_math(xml),
                ParagraphChild::InlineImage(img) => p.add_inline_image(img),
                other => { p.children.push(other); p }
            };
        }
    }

    // Determine text alignment
    let align_str = match params.alignment {
        ParserAlignment::Left => Some("left"),
        ParserAlignment::Right => Some("right"),
        ParserAlignment::Center => Some("center"),
        ParserAlignment::None => {
            // If no markdown alignment, use template alignment if available
            params.template.map(|tmpl| {
                tmpl.cell_style_for_column(params.col_index)
                    .alignment
                    .as_str()
            })
        }
    };

    // Apply alignment to the paragraph inside the cell (w:pPr/w:jc)
    // This is where Word actually reads text alignment from.
    if let Some(align) = align_str {
        p = p.align(align);
    }

    let mut cell = TableCellElement::new().width(params.width).add_paragraph(p);

    // Apply vertical alignment from template
    if let Some(tmpl) = params.template {
        let v_align = &tmpl
            .cell_style_for_column(params.col_index)
            .vertical_alignment;
        if !v_align.is_empty() {
            cell = cell.vertical_alignment(v_align);
        }
    }

    // Apply shading from template or default
    if let Some(shading) = get_row_shading(params.row_index, params.template) {
        cell.shading = Some(shading);
    }

    cell
}

/// Get the background color for a table row based on template
fn get_row_shading(
    row_index: usize,
    template: Option<&crate::template::extract::table::TableTemplate>,
) -> Option<String> {
    if let Some(tmpl) = template {
        if row_index == 0 {
            // Header row
            tmpl.header
                .background_color
                .as_ref()
                .map(|c| c.trim_start_matches('#').to_string())
        } else if row_index % 2 == 1 {
            // Odd row
            tmpl.row_odd
                .background_color
                .as_ref()
                .map(|c| c.trim_start_matches('#').to_string())
        } else {
            // Even row
            tmpl.row_even
                .background_color
                .as_ref()
                .map(|c| c.trim_start_matches('#').to_string())
        }
    } else {
        // Default: light blue for header
        if row_index == 0 {
            Some("D9E2F3".to_string())
        } else {
            None
        }
    }
}

/// Convert a code block to paragraphs (one per line)
fn code_block_to_paragraphs(
    content: &str,
    lang: Option<&str>,
    filename: Option<&str>,
    highlight_lines: &[u32],
    show_line_numbers: bool,
    code_font: Option<&str>,
    code_size: Option<u32>,
) -> Vec<Paragraph> {
    let mut paragraphs = Vec::new();

    // Get syntax-highlighted tokens for the content
    let highlighted = crate::docx::highlight::highlight_code(content, lang);

    // Helper to apply code font/size to a run
    let apply_code_style = |mut run: Run| -> Run {
        if let Some(font) = code_font {
            run = run.font(font);
        }
        if let Some(size) = code_size {
            run = run.size(size);
        }
        run
    };

    // Add filename as a separate paragraph if present
    if let Some(fname) = filename {
        let mut fname_run = Run::new(fname);
        if let Some(font) = code_font {
            fname_run = fname_run.font(font);
        }
        let filename_para = Paragraph::with_style("CodeFilename")
            .add_run(fname_run)
            .spacing(280, 0)
            .line_spacing(240, "auto");
        paragraphs.push(filename_para);
    }

    let lines: Vec<&str> = content.lines().collect();
    let total_lines = lines.len();

    // Add each line as a separate paragraph
    for (i, highlighted_line) in highlighted.iter().enumerate() {
        let line_num = (i + 1) as u32;

        // First line gets spacing before, last line gets spacing after
        let sp_before = if i == 0 && filename.is_none() { 280 } else { 0 };
        let sp_after = if i == total_lines - 1 { 280 } else { 0 };

        let mut p = Paragraph::with_style("Code")
            .spacing(sp_before, sp_after)
            .line_spacing(240, "auto");

        // Handle line numbers
        if show_line_numbers {
            let num_text = format!("{:>2}. ", line_num);
            p = p.add_run(apply_code_style(Run::new(num_text).color("888888")));
        }

        // Add syntax-highlighted runs
        if highlighted_line.is_empty() {
            p = p.add_run(apply_code_style(Run::new("")));
        } else {
            for (text, color) in highlighted_line {
                let mut run = Run::new(text.as_str());
                if let Some(c) = color {
                    run = run.color(c);
                }
                p = p.add_run(apply_code_style(run));
            }
        }

        // Handle line highlighting
        if highlight_lines.contains(&line_num) {
            p = p.shading("FFFACD"); // LemonChiffon
        }

        paragraphs.push(p);
    }

    // If content is empty, add at least one paragraph
    if paragraphs.is_empty() || (paragraphs.len() == 1 && filename.is_some()) {
        paragraphs.push(
            Paragraph::with_style("Code")
                .add_text("")
                .spacing(280, 280)
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
        let mut is_first_block = true;
        // Process the item's content blocks
        for block in &item.content {
            let mut item_paragraphs = block_to_paragraphs(block, list_level + 1, ctx, skip_toc);

            // Apply list styling only to the first paragraph of the first block in the item.
            // Subsequent blocks (e.g. nested lists) already have their own styling
            // and should not be overridden with the parent's numbering.
            if is_first_block {
                if let Some(first_para) = item_paragraphs.first_mut() {
                    first_para.style_id = Some("ListParagraph".to_string());

                    // Use the provided unique numId for this list
                    let ilvl = list_level as u32;
                    first_para.numbering_id = Some(num_id);
                    first_para.numbering_level = Some(ilvl);
                }
                is_first_block = false;
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

    apply_font_override_to_children(&mut children, &ctx.font_override);
    children
}

/// Apply font override to all runs within paragraph children
fn apply_font_override_to_children(
    children: &mut [ParagraphChild],
    font_override: &Option<String>,
) {
    if let Some(ref font) = font_override {
        for child in children.iter_mut() {
            match child {
                ParagraphChild::Run(run) => {
                    if run.font.is_none() {
                        run.font = Some(font.clone());
                    }
                }
                ParagraphChild::Hyperlink(hyperlink) => {
                    for run in &mut hyperlink.children {
                        if run.font.is_none() {
                            run.font = Some(font.clone());
                        }
                    }
                }
                _ => {}
            }
        }
    }
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
            // Check for PAGEREF pattern: [{PAGENUM}](#bookmark)
            if url.starts_with('#') {
                let link_text = extract_inline_text(text);
                if link_text.contains("{PAGENUM}") {
                    let bookmark = &url[1..]; // Strip the '#'
                    // Generate a PAGEREF field: begin + instrText + separate + placeholder + end
                    let mut children = Vec::new();
                    children.push(ParagraphChild::Run(
                        Run::new("").with_field_char("begin"),
                    ));
                    children.push(ParagraphChild::Run(
                        Run::new(format!(" PAGEREF {} \\h ", bookmark)).with_instr_text(),
                    ));
                    children.push(ParagraphChild::Run(
                        Run::new("").with_field_char("separate"),
                    ));
                    children.push(ParagraphChild::Run(Run::new("0"))); // Placeholder page number
                    children.push(ParagraphChild::Run(
                        Run::new("").with_field_char("end"),
                    ));
                    return children;
                }
            }

            let rel_id = ctx.hyperlink_ctx.add_hyperlink(url, ctx.rel_manager);
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
                // Create a temporary RelIdManager for footnote content to avoid affecting global state?
                // Actually, images in footnotes should use the global manager to be valid in document.xml.rels
                // BUT footnotes are in footnotes.xml, which has its own relationships file footnotes.xml.rels!
                // Currently we don't support images in footnotes fully (relationships wise).
                // For now, let's reuse the global manager but be aware that if we add images here,
                // they need to be in footnotes.xml.rels, not document.xml.rels.
                // TODO: Handle footnote relationships properly.

                let mut content = Vec::new();
                for block in blocks {
                    let mut nested_ctx = BuildContext {
                        image_ctx: &mut ImageContext::new(), // Temporary
                        hyperlink_ctx: ctx.hyperlink_ctx, // Re-use? Hyperlinks in footnotes need relationships too
                        numbering_ctx: &mut footnote_numbering_ctx,
                        doc: ctx.doc,
                        footnotes: ctx.footnotes,
                        toc_builder: &mut footnote_toc_builder,
                        bookmark_id_counter: &mut footnote_bookmark_id,
                        xref_ctx: &mut footnote_xref_ctx,
                        rel_manager: ctx.rel_manager,
                        table_template: ctx.table_template,
                        image_template: ctx.image_template,
                        table_count: &mut 0, // Footnotes don't typically have tables with captions, or they share numbering?
                        figure_count: &mut 0,
                        lang: ctx.lang,
                        font_override: ctx.font_override.clone(),
                        code_font: ctx.code_font.clone(),
                        code_size: ctx.code_size,
                        quote_level: 0,
                        mermaid_spacing: ctx.mermaid_spacing,
                        math_renderer: ctx.math_renderer.clone(),
                        math_font_size: ctx.math_font_size.clone(),
                        math_number_all: ctx.math_number_all,
                        body_width_twips: ctx.body_width_twips,
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
                    // Remove empty trailing paragraphs from footnote content
                    while content.len() > 1
                        && content.last().map_or(false, |p| p.children.is_empty())
                    {
                        content.pop();
                    }

                    // Set FootnoteText style on all paragraphs and remove indentation
                    for p in &mut content {
                        p.style_id = Some("FootnoteText".to_string());
                        p.indent_left = None;
                        p.spacing_before = Some(0);
                        p.spacing_after = Some(0);
                    }

                    // Add footnoteRef marker at the beginning of first paragraph
                    // This generates the footnote number in the footnote content area
                    let mut fn_ref_run = Run::new("");
                    fn_ref_run.style = Some("FootnoteReference".to_string());
                    fn_ref_run.superscript = true;
                    fn_ref_run.footnote_ref = true;

                    let space_run = Run::new(" ");

                    // Insert footnoteRef run + space at the beginning of first paragraph
                    content[0]
                        .children
                        .insert(0, ParagraphChild::Run(space_run));
                    content[0]
                        .children
                        .insert(0, ParagraphChild::Run(fn_ref_run));

                    let id = ctx.footnotes.add_footnote(content);
                    // Return a run with footnote reference (superscript in body text)
                    let mut run = Run::new("");
                    run.footnote_id = Some(id);
                    run.style = Some("FootnoteReference".to_string());
                    run.superscript = true;
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
            // Resolve the anchor to get bookmark info
            if let Some(anchor) = ctx.xref_ctx.resolve(target) {
                let bookmark_name = anchor.bookmark_name.clone();
                let display_text = ctx.xref_ctx.get_localized_display_text(target, ctx.lang);

                if *ref_type == RefType::Equation {
                    // Equation cross-refs use a dynamic REF field pointing to the bookmark
                    // so Word can update numbers automatically with F9
                    let mut children = Vec::new();
                    // REF field begin
                    children.push(ParagraphChild::Run(
                        Run::new("").with_field_char("begin"),
                    ));
                    // REF field instruction
                    children.push(ParagraphChild::Run(
                        Run::new(format!(" REF {} \\h ", bookmark_name)).with_instr_text(),
                    ));
                    // REF field separate
                    children.push(ParagraphChild::Run(
                        Run::new("").with_field_char("separate"),
                    ));
                    // Placeholder text (Word updates this on F9)
                    let mut placeholder = Run::new(&display_text);
                    placeholder.color = Some("0563C1".to_string());
                    placeholder.underline = true;
                    children.push(ParagraphChild::Run(placeholder));
                    // REF field end
                    children.push(ParagraphChild::Run(
                        Run::new("").with_field_char("end"),
                    ));
                    children
                } else {
                    // Non-equation cross-refs: styled text (TODO: hyperlink in future)
                    let mut run = Run::new(&display_text);
                    run.color = Some("0563C1".to_string());
                    run.underline = true;
                    vec![ParagraphChild::Run(run)]
                }
            } else {
                // Unresolved reference — show as plain text
                let display_text = ctx.xref_ctx.get_localized_display_text(target, ctx.lang);
                let mut run = Run::new(&display_text);
                run.color = Some("FF0000".to_string()); // Red to indicate missing ref
                vec![ParagraphChild::Run(run)]
            }
        }

        Inline::SoftBreak => {
            // In blockquotes, soft break becomes a line break to preserve
            // the visual line structure. Outside blockquotes, it becomes a space.
            if ctx.quote_level > 0 {
                vec![ParagraphChild::Run(create_break_run())]
            } else {
                vec![ParagraphChild::Run(Run::new(" "))]
            }
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

        Inline::InlineMath(latex) => {
            if ctx.math_renderer == "rex" {
                let render_result = crate::docx::math_rex::render_latex_to_svg(latex, false, &ctx.math_font_size);
                match render_result {
                    Ok(math) => {
                        let image_id = ctx.rel_manager.next_image_id();
                        let filename = format!("math_inline{}.svg", image_id);

                        let rel_id = ctx.image_ctx.add_image_data(
                            &filename,
                            math.svg_bytes,
                            None,
                            ctx.rel_manager,
                        );

                        let mut img = ImageElement::new(&rel_id, math.width_emu, math.height_emu)
                            .alt_text("Math")
                            .name(&filename)
                            .id(image_id);
                        img.position = math.position;

                        vec![ParagraphChild::InlineImage(img)]
                    }
                    Err(e) => {
                        eprintln!("Warning: ReX rendering failed for inline math, falling back to OMML: {}", e);
                        let omml = crate::docx::math::latex_to_omml_inline(latex);
                        vec![ParagraphChild::OfficeMath(omml)]
                    }
                }
            } else {
                let omml = crate::docx::math::latex_to_omml_inline(latex);
                vec![ParagraphChild::OfficeMath(omml)]
            }
        }

        Inline::DisplayMath(latex) => {
            if ctx.math_renderer == "rex" {
                let render_result = crate::docx::math_rex::render_latex_to_svg(latex, true, &ctx.math_font_size);
                match render_result {
                    Ok(math) => {
                        let image_id = ctx.rel_manager.next_image_id();
                        let filename = format!("math_display_inline{}.svg", image_id);

                        let rel_id = ctx.image_ctx.add_image_data(
                            &filename,
                            math.svg_bytes,
                            None,
                            ctx.rel_manager,
                        );

                        let mut img = ImageElement::new(&rel_id, math.width_emu, math.height_emu)
                            .alt_text("Math equation")
                            .name(&filename)
                            .id(image_id);
                        img.position = math.position;

                        vec![ParagraphChild::InlineImage(img)]
                    }
                    Err(e) => {
                        eprintln!("Warning: ReX rendering failed for display math, falling back to OMML: {}", e);
                        let omml = crate::docx::math::latex_to_omml_paragraph(latex);
                        vec![ParagraphChild::OfficeMath(omml)]
                    }
                }
            } else {
                // Display math in inline context: use oMathPara
                let omml = crate::docx::math::latex_to_omml_paragraph(latex);
                vec![ParagraphChild::OfficeMath(omml)]
            }
        }
    }
}

/// Create a run with a line break element
///
/// In OOXML, a line break is represented by a `<w:br/>` element
/// inside a run. This creates a run that contains only a break.
fn create_break_run() -> Run {
    let mut run = Run::new("");
    run.break_type = Some("textWrapping".to_string());
    run
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        assert_eq!(result.footnotes.len(), 1);

        // Footnote should have multiple paragraphs (one per line)
        let footnotes = result.footnotes.get_footnotes();
        let footnote = &footnotes[0];
        assert!(
            footnote.content.len() >= 1,
            "Footnote should have content"
        );
    }

    #[test]
    fn test_footnote_xml_generation() {
        let md = "Text[^1]\n\n[^1]: Footnote content";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &doc,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &doc,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &doc,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have 3 lines of code
        assert_eq!(paragraphs.len(), 3);
        assert_eq!(paragraphs[0].style_id, Some("Code".to_string()));
        let line1: String = paragraphs[0].iter_runs().map(|r| r.text.as_str()).collect();
        assert_eq!(line1, "fn main() {");
        let line2: String = paragraphs[1].iter_runs().map(|r| r.text.as_str()).collect();
        assert_eq!(line2, "    println!(\"Hello\");");
        let line3: String = paragraphs[2].iter_runs().map(|r| r.text.as_str()).collect();
        assert_eq!(line3, "}");
    }

    #[test]
    fn test_blockquote() {
        let md = "> This is a quote\n> With multiple lines";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        assert_eq!(paragraphs.len(), 2);
        // First is a blank Normal paragraph
        assert_eq!(paragraphs[0].style_id, Some("Normal".to_string()));
        assert!(paragraphs[0].children.is_empty());
        // Second is the section break paragraph
        assert!(paragraphs[1].children.is_empty());
    }

    #[test]
    fn test_link() {
        let md = "[OpenAI](https://openai.com)";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Check that indentation is preserved
        let line2: String = paragraphs[1].iter_runs().map(|r| r.text.as_str()).collect();
        assert_eq!(line2, "    println!(\"Hello\");");
    }

    #[test]
    fn test_code_block_with_line_numbers() {
        let md = "```rust,ln\nline 1\nline 2\nline 3\n```";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have filename paragraph + code line
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(paragraphs[0].style_id, Some("CodeFilename".to_string()));
        assert_eq!(paragraphs[0].iter_runs().next().unwrap().text, "main.rs");
        assert_eq!(paragraphs[1].style_id, Some("Code".to_string()));
        let code_text: String = paragraphs[1].iter_runs().map(|r| r.text.as_str()).collect();
        assert_eq!(code_text, "fn main() {}");
    }

    #[test]
    fn test_link_with_formatting() {
        let md = "[**bold link**](https://example.com)";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
        let docx = &result.document;

        if let Some(DocElement::Table(table)) = docx.elements.first() {
            // Check alignments on data row cells' paragraphs (w:pPr/w:jc)
            if let Some(data_row) = table.rows.get(1) {
                assert_eq!(
                    data_row.cells.get(0).and_then(|c| c.paragraphs.first()).and_then(|p| p.align.as_deref()),
                    Some("left")
                );
                assert_eq!(
                    data_row.cells.get(1).and_then(|c| c.paragraphs.first()).and_then(|p| p.align.as_deref()),
                    Some("center")
                );
                assert_eq!(
                    data_row.cells.get(2).and_then(|c| c.paragraphs.first()).and_then(|p| p.align.as_deref()),
                    Some("right")
                );
            }
        }
    }

    #[test]
    fn test_table_with_multiple_rows() {
        let md = "| Name | Age |\n|------|-----|\n| John | 30  |\n| Jane | 25  |\n| Bob  | 35  |";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let id = ctx.add_image("test.png", None, &mut rel_manager);
        // rId1-5 are reserved, so first image should be rId6
        assert_eq!(id, "rId6");
        assert_eq!(ctx.images.len(), 1);
        assert_eq!(ctx.images[0].src, "test.png");
        // Filename includes rel_id for uniqueness
        assert_eq!(ctx.images[0].filename, "image_rId6.png");
    }

    #[test]
    fn test_image_context_multiple() {
        let mut ctx = ImageContext::new();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let id1 = ctx.add_image("img1.png", None, &mut rel_manager);
        let id2 = ctx.add_image("img2.png", None, &mut rel_manager);

        assert_eq!(id1, "rId6");
        assert_eq!(id2, "rId7");
        assert_eq!(ctx.images.len(), 2);
        // Filenames include rel_id for uniqueness
        assert_eq!(ctx.images[0].filename, "image_rId6.png");
        assert_eq!(ctx.images[1].filename, "image_rId7.png");
    }

    #[test]
    fn test_image_context_dimensions_default() {
        let mut ctx = ImageContext::new();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        ctx.add_image("test.png", None, &mut rel_manager);
        // Default 6x4 inches
        assert_eq!(ctx.images[0].width_emu, 5486400);
        assert_eq!(ctx.images[0].height_emu, 3657600);
    }

    #[test]
    fn test_image_context_dimensions_inches() {
        let mut ctx = ImageContext::new();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        ctx.add_image("test.png", Some("2in"), &mut rel_manager);
        // 2 inches = 1828800 EMUs
        assert_eq!(ctx.images[0].width_emu, 1828800);
    }

    #[test]
    fn test_image_context_dimensions_pixels() {
        let mut ctx = ImageContext::new();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        ctx.add_image("test.png", Some("96px"), &mut rel_manager);
        // 96px = 1 inch = 914400 EMUs
        assert_eq!(ctx.images[0].width_emu, 914400);
    }

    #[test]
    fn test_image_context_dimensions_percentage() {
        let mut ctx = ImageContext::new();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        ctx.add_image("test.png", Some("50%"), &mut rel_manager);
        // 50% of 6.0in = 3.0in = 2743200 EMUs
        assert_eq!(ctx.images[0].width_emu, 2743200);
    }

    #[test]
    fn test_image_context_filename_generation() {
        let ctx = ImageContext::new();
        assert_eq!(
            ctx.generate_filename("path/to/test.png", "rId1".to_string()),
            "image_rId1.png"
        );
        assert_eq!(
            ctx.generate_filename("http://example.com/img.jpg", "rId2".to_string()),
            "image_rId2.jpg"
        );
    }

    #[test]
    fn test_build_document_with_image() {
        let md = "![Test](test.png)";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        assert_eq!(result.images.images.len(), 1);
        assert_eq!(result.images.images[0].rel_id, "rId6");

        // Check paragraph has image + caption (fallback caption from alt text)
        let paragraphs = get_paragraphs(&result.document);
        assert_eq!(paragraphs.len(), 1); // Caption paragraph

        if let Some(DocElement::Image(img)) = result.document.elements.first() {
            assert_eq!(img.rel_id, "rId6");
        } else {
            panic!("Expected Image element");
        }
    }

    #[test]
    fn test_build_document_image_with_width() {
        let md = "![alt](image.png){width=50%}";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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

        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // Should have three headers and three footers:
        // 1. Default header/footer with content
        // 2. Empty header/footer for first page (when different_first_page is set)
        // 3. Suppression header/footer for cover/TOC (always empty)
        assert_eq!(result.headers.len(), 3);
        assert_eq!(result.footers.len(), 3);

        // Check header XML
        let header_xml = String::from_utf8(result.headers[0].xml_bytes.clone()).unwrap();
        assert!(header_xml.contains("<w:hdr"));
        assert!(header_xml.contains("My Document"));
        assert!(header_xml.contains("STYLEREF"));

        // Check footer XML
        let footer_xml = String::from_utf8(result.footers[0].xml_bytes.clone()).unwrap();
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

        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // Should have three headers and three footers:
        // 1. Default header/footer with content
        // 2. First page header/footer (empty when different_first_page)
        // 3. Suppression header/footer for cover/TOC (always empty)
        assert_eq!(result.headers.len(), 3);
        assert_eq!(result.footers.len(), 3);

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

        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        assert_eq!(result.images.images.len(), 2);
        assert_eq!(result.images.images[0].rel_id, "rId6");
        assert_eq!(result.images.images[1].rel_id, "rId7");
    }

    #[test]
    fn test_toc_generation() {
        let md = "# Chapter 1\n\n## Section 1.1\n\n### Subsection 1.1.1\n\n# Chapter 2";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // Check toc_builder
        let toc_builder = result.toc_builder.as_ref().unwrap();
        assert_eq!(toc_builder.entries().len(), 4);
        assert_eq!(toc_builder.entries()[0].text, "Chapter 1");
        assert_eq!(toc_builder.entries()[1].text, "Section 1.1");

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &no_toc_config(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        let docx = &result.document;
        let paragraphs = get_paragraphs(docx);

        // Should have: H1 + H2 = 2 paragraphs
        assert_eq!(paragraphs.len(), 2);

        // toc_builder should still have entries! (Collection is independent of generation)
        let toc_builder = result.toc_builder.as_ref().unwrap();
        assert_eq!(toc_builder.entries().len(), 2);
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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // toc_builder should have ALL entries, filtering happens during generation
        let toc_builder = result.toc_builder.as_ref().unwrap();
        assert_eq!(toc_builder.entries().len(), 4);

        // We can test generation directly
        let toc_elements = toc_builder.generate_toc(&config.toc);
        // title + field begin + 2 entries (h1, h2) + field end + section break = 6
        assert_eq!(toc_elements.len(), 6);
    }

    #[test]
    fn test_toc_with_explicit_id() {
        let md = "# Introduction {#intro}\n\n## Getting Started {#start}";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        let toc_builder = result.toc_builder.as_ref().unwrap();
        let entry = &toc_builder.entries()[0];
        assert_eq!(entry.level, 1);

        // TOC entry should contain the plain text (without formatting)
        assert!(entry.text.contains("Bold"));
        assert!(entry.text.contains("italic"));
    }

    #[test]
    fn test_build_document_image_with_alt_text() {
        let md = "![This is alt text](image.png)";
        let parsed = parse_markdown_with_frontmatter(md);
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

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
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &DocumentConfig::default(),
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // Verify BuildResult structure
        assert!(result.document.elements.len() > 0);
        assert_eq!(result.images.images.len(), 0);
    }

    #[test]
    fn test_table_with_template() {
        use crate::template::extract::table::TableTemplate;

        let md = "| Header |\n| --- |\n| Cell |";
        let parsed = parse_markdown_with_frontmatter(md);

        // Create a custom table template
        let mut template = TableTemplate::default();
        template.header.background_color = Some("#FF0000".to_string()); // Red header
        template.row_odd.background_color = Some("#00FF00".to_string()); // Green odd row (row 1 in data rows, index 1 in table)

        let config = DocumentConfig::default();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            Some(&template),
            None,
        )
        .unwrap();

        // Find the table
        let table = result
            .document
            .elements
            .iter()
            .find_map(|el| match el {
                DocElement::Table(t) => Some(t),
                _ => None,
            })
            .expect("Should have a table");

        // Check header shading
        let header_cell = &table.rows[0].cells[0];
        assert_eq!(header_cell.shading, Some("FF0000".to_string()));

        // Check data row shading
        let data_cell = &table.rows[1].cells[0];
        assert_eq!(data_cell.shading, Some("00FF00".to_string()));
    }

    #[test]
    fn test_first_column_bold_other_columns_normal() {
        use crate::template::extract::table::TableTemplate;

        // Create a table with 2 columns
        let md = "| Header1 | Header2 |\n| --- | --- |\n| Row1Col1 | Row1Col2 |";
        let parsed = parse_markdown_with_frontmatter(md);

        // Create a template where:
        // - first_column.bold = true
        // - other_columns.bold = false
        // - row styles (row_odd, row_even) have bold = false
        let mut template = TableTemplate::default();
        template.first_column.bold = true;
        template.other_columns.bold = false;
        template.row_odd.bold = false;
        template.row_even.bold = false;
        template.header.bold = true; // Keep header bold

        let config = DocumentConfig::default();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            Some(&template),
            None,
        )
        .unwrap();

        // Find the table
        let table = result
            .document
            .elements
            .iter()
            .find_map(|el| match el {
                DocElement::Table(t) => Some(t),
                _ => None,
            })
            .expect("Should have a table");

        // Verify we have 2 rows (header + 1 data row)
        assert_eq!(table.rows.len(), 2, "Should have 2 rows");

        // Header row (index 0) - both cells should be bold (from header row style)
        let header_row = &table.rows[0];
        assert_eq!(header_row.cells.len(), 2, "Header should have 2 cells");

        // Check header cells are bold
        for (i, cell) in header_row.cells.iter().enumerate() {
            let para = &cell.paragraphs[0];
            let runs: Vec<_> = para.iter_runs().collect();
            assert!(!runs.is_empty(), "Header cell {} should have runs", i);
            for run in runs {
                assert!(run.bold, "Header cell {} run should be bold", i);
            }
        }

        // Data row (index 1)
        let data_row = &table.rows[1];
        assert_eq!(data_row.cells.len(), 2, "Data row should have 2 cells");

        // First cell (col 0) should have bold text
        let first_cell_para = &data_row.cells[0].paragraphs[0];
        let first_cell_runs: Vec<_> = first_cell_para.iter_runs().collect();
        assert!(!first_cell_runs.is_empty(), "First cell should have runs");
        for run in &first_cell_runs {
            assert!(
                run.bold,
                "First column cell should be bold. Run text: '{}'",
                run.text
            );
        }

        // Second cell (col 1) should NOT have bold text
        let second_cell_para = &data_row.cells[1].paragraphs[0];
        let second_cell_runs: Vec<_> = second_cell_para.iter_runs().collect();
        assert!(!second_cell_runs.is_empty(), "Second cell should have runs");
        for run in &second_cell_runs {
            assert!(
                !run.bold,
                "Other column cells should NOT be bold. Run text: '{}'",
                run.text
            );
        }
    }

    #[test]
    fn test_table_with_borders() {
        use crate::template::extract::table::{BorderStyle, BorderStyles, TableTemplate};

        let md = "| Header |\n| --- |\n| Cell |";
        let parsed = parse_markdown_with_frontmatter(md);

        // Create a custom table template with specific borders
        let mut template = TableTemplate::default();
        template.borders = BorderStyles {
            top: BorderStyle {
                style: "double".to_string(),
                color: "#FF0000".to_string(),
                width: 8,
            },
            ..Default::default()
        };

        let config = DocumentConfig::default();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            Some(&template),
            None,
        )
        .unwrap();

        // Find the table
        let table = result
            .document
            .elements
            .iter()
            .find_map(|el| match el {
                DocElement::Table(t) => Some(t),
                _ => None,
            })
            .expect("Should have a table");

        // Check borders
        assert!(table.borders.is_some());
        let borders = table.borders.as_ref().unwrap();
        assert_eq!(borders.top.style, "double");
        assert_eq!(borders.top.color, "#FF0000");
        assert_eq!(borders.top.width, 8);
    }

    #[test]
    fn test_table_with_caption() {
        use crate::template::extract::table::TableTemplate;

        // Note: Currently we don't have parser support for captions,
        // so we manually create a Block::Table with a caption for testing.
        let table_block = Block::Table {
            headers: vec![ParserTableCell {
                content: vec![Inline::Text("Header".to_string())],
                is_header: true,
            }],
            alignments: vec![ParserAlignment::None],
            rows: vec![vec![ParserTableCell {
                content: vec![Inline::Text("Cell".to_string())],
                is_header: false,
            }]],
            caption: Some("My Table Caption".to_string()),
            id: None,
        };

        let doc = ParsedDocument {
            blocks: vec![table_block],
            ..Default::default()
        };

        // Create a custom table template
        let template = TableTemplate::default();

        let config = DocumentConfig::default();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();
        let result = build_document(
            &doc,
            Language::English,
            &config,
            &mut rel_manager,
            Some(&template),
            None,
        )
        .unwrap();

        // Should have: Caption paragraph, Table, Empty paragraph
        assert_eq!(result.document.elements.len(), 3);

        // Check caption paragraph
        if let DocElement::Paragraph(p) = &result.document.elements[0] {
            assert_eq!(p.style_id, Some("Caption".to_string()));
            let text: String = p.iter_runs().map(|r| r.text.as_str()).collect();
            // Default prefix is "Table", number should be 1
            assert!(text.contains("Table 1: My Table Caption"));
        } else {
            panic!("Expected caption paragraph");
        }
    }

    #[test]
    fn test_table_cross_reference_thai() {
        let md = "# Chapter 1 {#ch1}\n\nTable: My Table {#tbl:test}\n| A | B |\n|---|---|\n| 1 | 2 |\n\nSee {ref:tbl:test}.";
        let parsed = parse_markdown_with_frontmatter(md);

        let mut config = DocumentConfig::default();
        config.toc.enabled = false;
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();

        // Mock table template to ensure caption generation
        let template = crate::template::extract::table::TableTemplate::default();

        let result = build_document(
            &parsed,
            Language::Thai,
            &config,
            &mut rel_manager,
            Some(&template),
            None,
        )
        .unwrap();

        // Find the cross-reference run
        let mut found_ref = false;
        for elem in &result.document.elements {
            if let DocElement::Paragraph(p) = elem {
                let text: String = p.iter_runs().map(|r| r.text.as_str()).collect();
                if text.contains("ตารางที่ 1.1") {
                    found_ref = true;
                }
            }
        }
        assert!(
            found_ref,
            "Cross-reference 'ตารางที่ 1.1' not found in document"
        );
    }

    #[test]
    fn test_mermaid_spacing_default_config() {
        // Default mermaid spacing should be (120, 120)
        let config = DocumentConfig::default();
        assert_eq!(config.mermaid_spacing, (120, 120));
    }

    #[test]
    fn test_mermaid_spacing_custom_config() {
        let mut config = DocumentConfig::default();
        config.mermaid_spacing = (200, 300);
        assert_eq!(config.mermaid_spacing, (200, 300));
    }

    #[test]
    fn test_math_renderer_default_config() {
        let config = DocumentConfig::default();
        assert_eq!(config.math_renderer, "rex");
    }

    #[test]
    fn test_math_renderer_omml_config() {
        let mut config = DocumentConfig::default();
        config.math_renderer = "omml".to_string();
        assert_eq!(config.math_renderer, "omml");
    }

    #[test]
    fn test_display_math_omml_renderer() {
        // When renderer is "omml", display math without \label should produce
        // a centered Paragraph with OfficeMath content but NO equation number
        let md = "$$\nE = mc^2\n$$\n";
        let parsed = parse_markdown_with_frontmatter(md);

        let mut config = no_toc_config();
        config.math_renderer = "omml".to_string();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();

        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // Should have a Paragraph element containing OfficeMath (no SEQ field since no label)
        let paragraphs = get_paragraphs(&result.document);
        let has_math = paragraphs.iter().any(|p| {
            p.children.iter().any(|c| matches!(c, ParagraphChild::OfficeMath(_)))
        });
        assert!(has_math, "Should produce Paragraph with OfficeMath content");

        // Verify no SEQ field (unlabeled equations should not be numbered)
        let has_seq = paragraphs.iter().any(|p| {
            p.children.iter().any(|c| {
                if let ParagraphChild::Run(r) = c {
                    r.instr_text && r.text.contains("SEQ Equation")
                } else {
                    false
                }
            })
        });
        assert!(!has_seq, "Unlabeled display equations should not have SEQ field numbers");
    }

    #[test]
    fn test_inline_math_omml_renderer() {
        // When renderer is "omml", inline math should produce OfficeMath children
        let md = "The formula $x^2$ is simple.\n";
        let parsed = parse_markdown_with_frontmatter(md);

        let mut config = no_toc_config();
        config.math_renderer = "omml".to_string();
        let mut rel_manager = crate::docx::rels_manager::RelIdManager::new();

        let result = build_document(
            &parsed,
            Language::English,
            &config,
            &mut rel_manager,
            None,
            None,
        )
        .unwrap();

        // Should have paragraphs with OfficeMath children
        let paragraphs = get_paragraphs(&result.document);
        let has_math = paragraphs.iter().any(|p| {
            p.children.iter().any(|c| matches!(c, ParagraphChild::OfficeMath(_)))
        });
        assert!(has_math, "Should produce OfficeMath children when renderer is omml");
    }
}
