//! Generate word/document.xml for DOCX

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;
use crate::i18n::detection::detect_language;
use crate::template::extract::table::{BorderStyle, BorderStyles, CellMargins};

/// Header and footer references for a section
#[derive(Debug, Clone, Default)]
pub struct HeaderFooterRefs {
    pub default_header_id: Option<String>, // rId for default header
    pub first_header_id: Option<String>,   // rId for first page header (can be empty)
    pub default_footer_id: Option<String>, // rId for default footer
    pub first_footer_id: Option<String>,   // rId for first page footer
    pub different_first_page: bool,        // Enable different first page
}

/// Text run with formatting
#[derive(Debug, Clone)]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub style: Option<String>,     // Character style
    pub font: Option<String>,      // Specific font override
    pub size: Option<u32>,         // Size in half-points
    pub color: Option<String>,     // Hex color (without #)
    pub highlight: Option<String>, // Highlight color
    pub preserve_space: bool,
    pub footnote_id: Option<i32>, // Footnote reference ID (if this is a footnote reference)
    pub field_char: Option<String>, // Field character type: "begin", "separate", "end"
    pub instr_text: bool,         // If true, this is instruction text for a field
    pub tab: bool,                // If true, this run contains a tab character
    pub lang: Option<String>,     // Language for spell-check (auto-detected from text)
    pub break_type: Option<String>, // "page", "column", "textWrapping"
}

impl Run {
    pub fn new(text: impl Into<String>) -> Self {
        let text_str = text.into();
        // Auto-detect language from text content for proper spell-checking
        let lang = Some(detect_language(&text_str).to_string());
        Self {
            text: text_str,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            style: None,
            font: None,
            size: None,
            color: None,
            highlight: None,
            preserve_space: true,
            footnote_id: None,
            field_char: None,
            instr_text: false,
            tab: false,
            lang,
            break_type: None,
        }
    }

    /// Set bold formatting
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }

    /// Set italic formatting
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }

    /// Set underline formatting
    pub fn underline(mut self) -> Self {
        self.underline = true;
        self
    }

    /// Set strikethrough formatting
    pub fn strike(mut self) -> Self {
        self.strike = true;
        self
    }

    /// Set character style
    pub fn style(mut self, style_id: &str) -> Self {
        self.style = Some(style_id.to_string());
        self
    }

    /// Set font override
    pub fn font(mut self, font: &str) -> Self {
        self.font = Some(font.to_string());
        self
    }

    /// Set font size in half-points
    pub fn size(mut self, size: u32) -> Self {
        self.size = Some(size);
        self
    }

    /// Set text color (hex without #)
    pub fn color(mut self, color: &str) -> Self {
        self.color = Some(color.to_string());
        self
    }

    /// Set highlight color
    pub fn highlight(mut self, color: &str) -> Self {
        self.highlight = Some(color.to_string());
        self
    }

    /// Set whether to preserve whitespace
    pub fn preserve_space(mut self, preserve: bool) -> Self {
        self.preserve_space = preserve;
        self
    }

    /// Set field character type (for TOC fields)
    pub fn with_field_char(mut self, char_type: &str) -> Self {
        self.field_char = Some(char_type.to_string());
        self
    }

    /// Mark this run as instruction text for a field
    pub fn with_instr_text(mut self) -> Self {
        self.instr_text = true;
        self
    }

    /// Add a tab character to this run
    pub fn with_tab(mut self) -> Self {
        self.tab = true;
        self
    }

    /// Add a page break to this run
    pub fn with_page_break(mut self) -> Self {
        self.break_type = Some("page".to_string());
        self
    }

    /// Write run XML to a writer
    pub fn write_xml<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;

        // Run properties
        if self.bold
            || self.italic
            || self.underline
            || self.strike
            || self.style.is_some()
            || self.font.is_some()
            || self.size.is_some()
            || self.color.is_some()
            || self.highlight.is_some()
            || self.footnote_id.is_some()
        {
            writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

            // ECMA-376 STRICT ORDERING for w:rPr:
            // 1. w:rStyle
            // 2. w:rFonts
            // 3. w:b
            // 4. w:i
            // 5. w:strike
            // 6. w:u (underline)
            // 7. w:color
            // 8. w:sz
            // 9. w:szCs
            // 10. w:highlight
            // 11. w:lang
            // 12. w14:ligatures

            // 1. Style
            if let Some(style) = &self.style {
                let mut elem = BytesStart::new("w:rStyle");
                elem.push_attribute(("w:val", style.as_str()));
                writer.write_event(Event::Empty(elem))?;
            }

            // 2. Font
            if let Some(font) = &self.font {
                let mut fonts = BytesStart::new("w:rFonts");
                fonts.push_attribute(("w:ascii", font.as_str()));
                fonts.push_attribute(("w:hAnsi", font.as_str()));
                fonts.push_attribute(("w:cs", font.as_str()));
                writer.write_event(Event::Empty(fonts))?;
            }

            // 3. Bold
            if self.bold {
                writer.write_event(Event::Empty(BytesStart::new("w:b")))?;
            }

            // 4. Italic
            if self.italic {
                writer.write_event(Event::Empty(BytesStart::new("w:i")))?;
            }

            // 5. Strikethrough
            if self.strike {
                writer.write_event(Event::Empty(BytesStart::new("w:strike")))?;
            }

            // 6. Underline
            if self.underline {
                let mut u = BytesStart::new("w:u");
                u.push_attribute(("w:val", "single"));
                writer.write_event(Event::Empty(u))?;
            }

            // 7. Color
            if let Some(color) = &self.color {
                let mut c = BytesStart::new("w:color");
                c.push_attribute(("w:val", color.as_str()));
                writer.write_event(Event::Empty(c))?;
            }

            // 8. Size
            if let Some(size) = self.size {
                let mut sz = BytesStart::new("w:sz");
                sz.push_attribute(("w:val", size.to_string().as_str()));
                writer.write_event(Event::Empty(sz))?;
            }

            // 9. Complex script size
            if let Some(size) = self.size {
                let mut sz_cs = BytesStart::new("w:szCs");
                sz_cs.push_attribute(("w:val", size.to_string().as_str()));
                writer.write_event(Event::Empty(sz_cs))?;
            }

            // 10. Highlight
            if let Some(highlight) = &self.highlight {
                let mut h = BytesStart::new("w:highlight");
                h.push_attribute(("w:val", highlight.as_str()));
                writer.write_event(Event::Empty(h))?;
            }

            // 11. Language setting - use auto-detected language for proper spell-checking
            let mut lang_elem = BytesStart::new("w:lang");
            let primary_lang = self.lang.as_deref().unwrap_or("en-US");
            lang_elem.push_attribute(("w:val", primary_lang));
            // For Thai text, set eastAsia and bidi to Thai; for others, keep Thai as fallback
            if primary_lang == "th-TH" {
                lang_elem.push_attribute(("w:eastAsia", "th-TH"));
                lang_elem.push_attribute(("w:bidi", "th-TH"));
            } else {
                // Keep Thai as eastAsia/bidi for mixed content support
                lang_elem.push_attribute(("w:eastAsia", "th-TH"));
                lang_elem.push_attribute(("w:bidi", "th-TH"));
            }
            writer.write_event(Event::Empty(lang_elem))?;

            // 12. Ligatures (Thai ligature support)
            let mut ligatures = BytesStart::new("w14:ligatures");
            ligatures.push_attribute(("w14:val", "all"));
            writer.write_event(Event::Empty(ligatures))?;

            writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        }

        // Field character (for TOC fields)
        if let Some(char_type) = &self.field_char {
            let mut fc = BytesStart::new("w:fldChar");
            fc.push_attribute(("w:fldCharType", char_type.as_str()));
            writer.write_event(Event::Empty(fc))?;
        }

        // Instruction text (for field codes)
        if self.instr_text && !self.text.is_empty() {
            let mut it = BytesStart::new("w:instrText");
            it.push_attribute(("xml:space", "preserve"));
            writer.write_event(Event::Start(it))?;
            writer.write_event(Event::Text(BytesText::new(&self.text)))?;
            writer.write_event(Event::End(BytesEnd::new("w:instrText")))?;
        }

        // Tab character
        if self.tab {
            writer.write_event(Event::Empty(BytesStart::new("w:tab")))?;
        }

        // Break
        if let Some(break_type) = &self.break_type {
            let mut br = BytesStart::new("w:br");
            if break_type != "textWrapping" {
                br.push_attribute(("w:type", break_type.as_str()));
            }
            writer.write_event(Event::Empty(br))?;
        }

        // Footnote reference (if present)
        if let Some(id) = self.footnote_id {
            let mut fn_ref = BytesStart::new("w:footnoteReference");
            fn_ref.push_attribute(("w:id", id.to_string().as_str()));
            writer.write_event(Event::Empty(fn_ref))?;
        }

        // Text (only if not instruction text and not empty)
        if !self.instr_text && !self.text.is_empty() {
            let mut t = BytesStart::new("w:t");
            if self.preserve_space {
                t.push_attribute(("xml:space", "preserve"));
            }
            writer.write_event(Event::Start(t))?;
            writer.write_event(Event::Text(BytesText::new(&self.text)))?;
            writer.write_event(Event::End(BytesEnd::new("w:t")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        Ok(())
    }
}

impl Default for Run {
    fn default() -> Self {
        Self::new("")
    }
}

/// Bookmark start element
#[derive(Debug, Clone)]
pub struct BookmarkStart {
    pub id: u32,      // Unique numeric ID
    pub name: String, // Bookmark name (e.g., "_Toc1_Introduction")
}

/// Hyperlink element for paragraphs
#[derive(Debug, Clone)]
pub struct Hyperlink {
    pub id: String,         // Relationship ID (rId...)
    pub children: Vec<Run>, // Hyperlinks usually contain runs
}

impl Hyperlink {
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            children: Vec::new(),
        }
    }

    pub fn add_run(mut self, run: Run) -> Self {
        self.children.push(run);
        self
    }
}

/// Child elements of a paragraph (Run or Hyperlink)
#[derive(Debug, Clone)]
pub enum ParagraphChild {
    Run(Run),
    Hyperlink(Hyperlink),
}

/// Paragraph with style and children (runs or hyperlinks)
#[derive(Debug, Clone)]
pub struct Paragraph {
    pub style_id: Option<String>,
    pub children: Vec<ParagraphChild>,
    pub numbering_id: Option<u32>,
    pub numbering_level: Option<u32>,
    pub align: Option<String>,       // "left", "center", "right", "both"
    pub spacing_before: Option<u32>, // In twips
    pub spacing_after: Option<u32>,  // In twips
    pub indent_left: Option<u32>,    // In twips
    pub line: Option<i32>,           // 240ths of a line (if auto) or twips
    pub line_rule: Option<String>,   // "auto", "exact", "atLeast"
    pub keep_with_next: bool,
    pub page_break_before: bool,
    pub shading: Option<String>,       // Fill color (hex without #)
    pub section_break: Option<String>, // "nextPage", "continuous", "evenPage", "oddPage"
    pub page_num_start: Option<u32>,   // Page number to restart at for section break
    pub suppress_header_footer: bool,  // Suppress header/footer references in sectPr
    pub empty_header_footer_refs: Option<HeaderFooterRefs>, // Empty header/footer refs to use when suppressing
    pub bookmark_start: Option<BookmarkStart>,              // Bookmark start element
    pub bookmark_end: bool,                                 // If true, close bookmark after content
    // Page layout for section breaks (in twips)
    pub sect_page_width: Option<u32>,    // Page width for sectPr
    pub sect_page_height: Option<u32>,   // Page height for sectPr
    pub sect_margin_top: Option<u32>,    // Top margin for sectPr
    pub sect_margin_right: Option<u32>,  // Right margin for sectPr
    pub sect_margin_bottom: Option<u32>, // Bottom margin for sectPr
    pub sect_margin_left: Option<u32>,   // Left margin for sectPr
    pub sect_margin_header: Option<u32>, // Header margin for sectPr
    pub sect_margin_footer: Option<u32>, // Footer margin for sectPr
    pub sect_margin_gutter: Option<u32>, // Gutter margin for sectPr
}

impl Paragraph {
    pub fn new() -> Self {
        Self {
            style_id: None,
            children: Vec::new(),
            numbering_id: None,
            numbering_level: None,
            align: None,
            spacing_before: Some(0),
            spacing_after: Some(0),
            indent_left: None,
            line: Some(240),
            line_rule: Some("auto".to_string()),
            keep_with_next: false,
            page_break_before: false,
            shading: None,
            section_break: None,
            page_num_start: None,
            suppress_header_footer: false,
            empty_header_footer_refs: None,
            bookmark_start: None,
            bookmark_end: false,
            sect_page_width: None,
            sect_page_height: None,
            sect_margin_top: None,
            sect_margin_right: None,
            sect_margin_bottom: None,
            sect_margin_left: None,
            sect_margin_header: None,
            sect_margin_footer: None,
            sect_margin_gutter: None,
        }
    }

    /// Create paragraph with style
    pub fn with_style(style_id: &str) -> Self {
        let mut p = Self::new();
        p.style_id = Some(style_id.to_string());
        p
    }

    /// Add a run to the paragraph
    pub fn add_run(mut self, run: Run) -> Self {
        self.children.push(ParagraphChild::Run(run));
        self
    }

    /// Add text as a run
    pub fn add_text(self, text: impl Into<String>) -> Self {
        self.add_run(Run::new(text))
    }

    /// Add a hyperlink to the paragraph
    pub fn add_hyperlink(mut self, hyperlink: Hyperlink) -> Self {
        self.children.push(ParagraphChild::Hyperlink(hyperlink));
        self
    }

    /// Get an iterator over all runs in the paragraph (including those inside hyperlinks)
    pub fn iter_runs(&self) -> impl Iterator<Item = &Run> {
        self.children.iter().filter_map(|child| match child {
            ParagraphChild::Run(run) => Some(run),
            ParagraphChild::Hyperlink(link) => link.children.first(),
        })
    }

    /// Get all runs as a vector (including those inside hyperlinks)
    pub fn get_runs(&self) -> Vec<&Run> {
        self.iter_runs().collect()
    }

    /// Set numbering
    pub fn numbering(mut self, id: u32, level: u32) -> Self {
        self.numbering_id = Some(id);
        self.numbering_level = Some(level);
        self
    }

    /// Set alignment
    pub fn align(mut self, align: &str) -> Self {
        self.align = Some(align.to_string());
        self
    }

    /// Set spacing in twips
    pub fn spacing(mut self, before: u32, after: u32) -> Self {
        self.spacing_before = Some(before);
        self.spacing_after = Some(after);
        self
    }

    /// Set left indent in twips
    pub fn indent(mut self, left: u32) -> Self {
        self.indent_left = Some(left);
        self
    }

    /// Set line spacing
    ///
    /// * `line` - Value (e.g., 240 for single spacing if rule is auto)
    /// * `rule` - "auto" (default), "exact", or "atLeast"
    pub fn line_spacing(mut self, line: i32, rule: &str) -> Self {
        self.line = Some(line);
        self.line_rule = Some(rule.to_string());
        self
    }

    /// Keep paragraph with next paragraph
    pub fn keep_with_next(mut self) -> Self {
        self.keep_with_next = true;
        self
    }

    /// Force page break before paragraph
    pub fn page_break_before(mut self) -> Self {
        self.page_break_before = true;
        self
    }

    /// Add a page break as a run
    pub fn page_break(mut self) -> Self {
        self.children
            .push(ParagraphChild::Run(Run::new("").with_page_break()));
        self
    }

    /// Set paragraph shading color (hex without #)
    pub fn shading(mut self, color: &str) -> Self {
        self.shading = Some(color.to_string());
        self
    }

    /// Add a section break to this paragraph
    pub fn section_break(mut self, break_type: &str) -> Self {
        self.section_break = Some(break_type.to_string());
        self
    }

    /// Set page number restart for section break
    pub fn page_num_start(mut self, start: u32) -> Self {
        self.page_num_start = Some(start);
        self
    }

    /// Check if this paragraph has a section break
    pub fn is_section_break(&self) -> bool {
        self.section_break.is_some()
    }

    /// Suppress header/footer references in section break
    pub fn suppress_header_footer(mut self) -> Self {
        self.suppress_header_footer = true;
        self
    }

    /// Set empty header/footer refs to explicitly suppress inheritance
    /// When set, these refs (pointing to empty header/footer files) will be used
    /// instead of inheriting from the previous section
    pub fn with_empty_header_footer_refs(mut self, refs: HeaderFooterRefs) -> Self {
        self.empty_header_footer_refs = Some(refs);
        self.suppress_header_footer = true; // Also set suppress flag
        self
    }

    /// Set page layout for section break (in twips)
    pub fn with_page_layout(
        mut self,
        width: Option<u32>,
        height: Option<u32>,
        margin_top: Option<u32>,
        margin_right: Option<u32>,
        margin_bottom: Option<u32>,
        margin_left: Option<u32>,
        margin_header: Option<u32>,
        margin_footer: Option<u32>,
        margin_gutter: Option<u32>,
    ) -> Self {
        self.sect_page_width = width;
        self.sect_page_height = height;
        self.sect_margin_top = margin_top;
        self.sect_margin_right = margin_right;
        self.sect_margin_bottom = margin_bottom;
        self.sect_margin_left = margin_left;
        self.sect_margin_header = margin_header;
        self.sect_margin_footer = margin_footer;
        self.sect_margin_gutter = margin_gutter;
        self
    }

    /// Wrap paragraph content with a bookmark
    pub fn with_bookmark(mut self, id: u32, name: &str) -> Self {
        self.bookmark_start = Some(BookmarkStart {
            id,
            name: name.to_string(),
        });
        self.bookmark_end = true;
        self
    }

    /// Write paragraph XML to a writer
    pub fn write_xml<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        header_footer_refs: Option<&HeaderFooterRefs>,
    ) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:p")))?;

        // Paragraph properties
        if self.style_id.is_some()
            || self.numbering_id.is_some()
            || self.align.is_some()
            || self.spacing_before.is_some()
            || self.spacing_after.is_some()
            || self.indent_left.is_some()
            || self.keep_with_next
            || self.page_break_before
            || self.shading.is_some()
            || self.section_break.is_some()
        {
            writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

            // ECMA-376 STRICT ORDERING for w:pPr:
            // 1. w:pStyle
            // 2. w:keepNext
            // 3. w:pageBreakBefore
            // 4. w:numPr
            // 5. w:pBdr (if any)
            // 6. w:shd
            // 7. w:tabs
            // 8. w:spacing
            // 9. w:ind
            // 10. w:jc (alignment)
            // 11. w:outlineLvl
            // 12. w:rPr (paragraph-level run properties, contains w14:ligatures)
            // 13. w:sectPr (must be last)

            // 1. Style
            if let Some(style) = &self.style_id {
                let mut elem = BytesStart::new("w:pStyle");
                elem.push_attribute(("w:val", style.as_str()));
                writer.write_event(Event::Empty(elem))?;
            }

            // 2. Keep with next
            if self.keep_with_next {
                writer.write_event(Event::Empty(BytesStart::new("w:keepNext")))?;
            }

            // 3. Page break before
            if self.page_break_before {
                writer.write_event(Event::Empty(BytesStart::new("w:pageBreakBefore")))?;
            }

            // 4. Numbering
            if let Some(num_id) = self.numbering_id {
                writer.write_event(Event::Start(BytesStart::new("w:numPr")))?;

                let mut ilvl = BytesStart::new("w:ilvl");
                ilvl.push_attribute((
                    "w:val",
                    self.numbering_level.unwrap_or(0).to_string().as_str(),
                ));
                writer.write_event(Event::Empty(ilvl))?;

                let mut num_id_elem = BytesStart::new("w:numId");
                num_id_elem.push_attribute(("w:val", num_id.to_string().as_str()));
                writer.write_event(Event::Empty(num_id_elem))?;

                writer.write_event(Event::End(BytesEnd::new("w:numPr")))?;
            }

            // 5. Paragraph border (not used in current implementation, placeholder for ordering)
            // 6. Shading
            if let Some(color) = &self.shading {
                let mut shd = BytesStart::new("w:shd");
                shd.push_attribute(("w:val", "clear"));
                shd.push_attribute(("w:color", "auto"));
                shd.push_attribute(("w:fill", color.as_str()));
                writer.write_event(Event::Empty(shd))?;
            }

            // 7. Tabs (not used in current implementation, placeholder for ordering)

            // 8. Spacing
            if self.spacing_before.is_some() || self.spacing_after.is_some() || self.line.is_some()
            {
                let mut spacing = BytesStart::new("w:spacing");
                if let Some(before) = self.spacing_before {
                    spacing.push_attribute(("w:before", before.to_string().as_str()));
                }
                if let Some(after) = self.spacing_after {
                    spacing.push_attribute(("w:after", after.to_string().as_str()));
                }
                if let Some(line) = self.line {
                    spacing.push_attribute(("w:line", line.to_string().as_str()));
                }
                if let Some(rule) = &self.line_rule {
                    spacing.push_attribute(("w:lineRule", rule.as_str()));
                }
                writer.write_event(Event::Empty(spacing))?;
            }

            // 9. Indent
            if let Some(indent) = self.indent_left {
                let mut indent_elem = BytesStart::new("w:ind");
                indent_elem.push_attribute(("w:left", indent.to_string().as_str()));
                writer.write_event(Event::Empty(indent_elem))?;
            }

            // 10. Alignment
            if let Some(align) = &self.align {
                let mut elem = BytesStart::new("w:jc");
                elem.push_attribute(("w:val", align.as_str()));
                writer.write_event(Event::Empty(elem))?;
            }

            // 11. Outline level (not used in current implementation, placeholder for ordering)

            // 12. Paragraph-level run properties with ligatures
            writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;
            // Add ligatures for Thai support
            let mut ligatures = BytesStart::new("w14:ligatures");
            ligatures.push_attribute(("w14:val", "all"));
            writer.write_event(Event::Empty(ligatures))?;
            writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;

            // 13. Section break (must be last in pPr)
            if let Some(break_type) = &self.section_break {
                writer.write_event(Event::Start(BytesStart::new("w:sectPr")))?;

                // Determine which refs to use:
                // 1. If empty_header_footer_refs is set, use those (for suppression via empty files)
                // 2. Otherwise, use passed-in refs (normal header/footer)
                let refs_to_use = self
                    .empty_header_footer_refs
                    .as_ref()
                    .or(header_footer_refs);

                // Write header/footer references
                if let Some(refs) = refs_to_use {
                    if let Some(ref id) = refs.default_header_id {
                        let mut header_ref = BytesStart::new("w:headerReference");
                        header_ref.push_attribute(("w:type", "default"));
                        header_ref.push_attribute(("r:id", id.as_str()));
                        writer.write_event(Event::Empty(header_ref))?;
                    }

                    // First page header (if different first page enabled)
                    if refs.different_first_page {
                        if let Some(ref id) = refs.first_header_id {
                            let mut header_ref = BytesStart::new("w:headerReference");
                            header_ref.push_attribute(("w:type", "first"));
                            header_ref.push_attribute(("r:id", id.as_str()));
                            writer.write_event(Event::Empty(header_ref))?;
                        }
                    }

                    // Default footer
                    if let Some(ref id) = refs.default_footer_id {
                        let mut footer_ref = BytesStart::new("w:footerReference");
                        footer_ref.push_attribute(("w:type", "default"));
                        footer_ref.push_attribute(("r:id", id.as_str()));
                        writer.write_event(Event::Empty(footer_ref))?;
                    }

                    // First page footer (if different first page enabled)
                    if refs.different_first_page {
                        if let Some(ref id) = refs.first_footer_id {
                            let mut footer_ref = BytesStart::new("w:footerReference");
                            footer_ref.push_attribute(("w:type", "first"));
                            footer_ref.push_attribute(("r:id", id.as_str()));
                            writer.write_event(Event::Empty(footer_ref))?;
                        }
                    }
                }

                let mut type_elem = BytesStart::new("w:type");
                type_elem.push_attribute(("w:val", break_type.as_str()));
                writer.write_event(Event::Empty(type_elem))?;

                // Page numbering restart
                if let Some(start) = self.page_num_start {
                    let mut pg_num_type = BytesStart::new("w:pgNumType");
                    pg_num_type.push_attribute(("w:start", start.to_string().as_str()));
                    writer.write_event(Event::Empty(pg_num_type))?;
                }

                // Page size (use configured values or defaults)
                let mut pg_sz = BytesStart::new("w:pgSz");
                pg_sz.push_attribute((
                    "w:w",
                    self.sect_page_width.unwrap_or(11906).to_string().as_str(),
                ));
                pg_sz.push_attribute((
                    "w:h",
                    self.sect_page_height.unwrap_or(16838).to_string().as_str(),
                ));
                writer.write_event(Event::Empty(pg_sz))?;

                // Margins (use configured values or defaults)
                let mut pg_mar = BytesStart::new("w:pgMar");
                pg_mar.push_attribute((
                    "w:top",
                    self.sect_margin_top.unwrap_or(1440).to_string().as_str(),
                ));
                pg_mar.push_attribute((
                    "w:right",
                    self.sect_margin_right.unwrap_or(1440).to_string().as_str(),
                ));
                pg_mar.push_attribute((
                    "w:bottom",
                    self.sect_margin_bottom.unwrap_or(1440).to_string().as_str(),
                ));
                pg_mar.push_attribute((
                    "w:left",
                    self.sect_margin_left.unwrap_or(1440).to_string().as_str(),
                ));
                pg_mar.push_attribute((
                    "w:header",
                    self.sect_margin_header.unwrap_or(708).to_string().as_str(),
                ));
                pg_mar.push_attribute((
                    "w:footer",
                    self.sect_margin_footer.unwrap_or(708).to_string().as_str(),
                ));
                pg_mar.push_attribute((
                    "w:gutter",
                    self.sect_margin_gutter.unwrap_or(0).to_string().as_str(),
                ));
                writer.write_event(Event::Empty(pg_mar))?;

                // Columns (single column by default)
                let mut cols = BytesStart::new("w:cols");
                cols.push_attribute(("w:space", "708"));
                writer.write_event(Event::Empty(cols))?;

                // Title page (different first page)
                if let Some(refs) = header_footer_refs {
                    if refs.different_first_page {
                        writer.write_event(Event::Empty(BytesStart::new("w:titlePg")))?;
                    }
                }

                // Document grid (for Asian text)
                let mut doc_grid = BytesStart::new("w:docGrid");
                doc_grid.push_attribute(("w:linePitch", "360"));
                writer.write_event(Event::Empty(doc_grid))?;

                writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
            }

            writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
        }

        // Bookmark start (if present)
        if let Some(ref bookmark) = self.bookmark_start {
            let mut bookmark_start = BytesStart::new("w:bookmarkStart");
            bookmark_start.push_attribute(("w:id", bookmark.id.to_string().as_str()));
            bookmark_start.push_attribute(("w:name", bookmark.name.as_str()));
            writer.write_event(Event::Empty(bookmark_start))?;
        }

        // Children (runs and hyperlinks)
        for child in &self.children {
            match child {
                ParagraphChild::Run(run) => {
                    run.write_xml(writer)?;
                }
                ParagraphChild::Hyperlink(hyperlink) => {
                    // Write <w:hyperlink r:id="...">
                    let mut link_elem = BytesStart::new("w:hyperlink");
                    link_elem.push_attribute(("r:id", hyperlink.id.as_str()));
                    writer.write_event(Event::Start(link_elem))?;

                    // Write hyperlink children (runs)
                    for run in &hyperlink.children {
                        run.write_xml(writer)?;
                    }

                    writer.write_event(Event::End(BytesEnd::new("w:hyperlink")))?;
                }
            }
        }

        // Bookmark end (if bookmark_end is true)
        if self.bookmark_end {
            if let Some(ref bookmark) = self.bookmark_start {
                let mut bookmark_end = BytesStart::new("w:bookmarkEnd");
                bookmark_end.push_attribute(("w:id", bookmark.id.to_string().as_str()));
                writer.write_event(Event::Empty(bookmark_end))?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:p")))?;
        Ok(())
    }
}

impl Default for Paragraph {
    fn default() -> Self {
        Self::new()
    }
}

/// Image element for embedding in document
#[derive(Debug, Clone)]
pub struct ImageElement {
    pub rel_id: String,   // Relationship ID (e.g., "rId4")
    pub width_emu: i64,   // Width in EMUs (914400 EMU = 1 inch)
    pub height_emu: i64,  // Height in EMUs
    pub alt_text: String, // Alt text / description
    pub name: String,     // Image name/filename
    pub id: u32,          // Unique ID for docPr
    pub border: Option<ImageBorderEffect>,
    pub shadow: Option<ImageShadowEffect>,
    pub effect_extent: Option<ImageEffectExtent>,
    pub alignment: Option<String>, // "left", "center", "right"
}

/// Image border effect for OOXML generation
#[derive(Debug, Clone)]
pub struct ImageBorderEffect {
    /// Fill type: "solid", "none"
    pub fill_type: String,
    /// Color value (hex without # or scheme name)
    pub color: String,
    /// Whether color is a scheme color (theme-based)
    pub is_scheme_color: bool,
    /// Border width in EMUs (None = default thin line)
    pub width: Option<u32>,
}

/// Image shadow effect for OOXML generation
#[derive(Debug, Clone)]
pub struct ImageShadowEffect {
    /// Blur radius in EMUs
    pub blur_radius: u32,
    /// Shadow distance in EMUs
    pub distance: u32,
    /// Direction in 60000ths of degree
    pub direction: u32,
    /// Alignment: "ctr", "tl", "tr", "bl", "br"
    pub alignment: String,
    /// Shadow color (hex without #)
    pub color: String,
    /// Opacity in thousandths (30000 = 30%)
    pub alpha: u32,
}

/// Effect extent for shadow/border space
#[derive(Debug, Clone, Default)]
pub struct ImageEffectExtent {
    pub left: u32,
    pub top: u32,
    pub right: u32,
    pub bottom: u32,
}

impl ImageElement {
    pub fn new(rel_id: &str, width_emu: i64, height_emu: i64) -> Self {
        Self {
            rel_id: rel_id.to_string(),
            width_emu,
            height_emu,
            alt_text: String::new(),
            name: String::new(),
            id: 1,
            border: None,
            shadow: None,
            effect_extent: None,
            alignment: None,
        }
    }

    pub fn alt_text(mut self, alt: &str) -> Self {
        self.alt_text = alt.to_string();
        self
    }

    pub fn name(mut self, name: &str) -> Self {
        self.name = name.to_string();
        self
    }

    pub fn id(mut self, id: u32) -> Self {
        self.id = id;
        self
    }

    pub fn with_border(mut self, border: ImageBorderEffect) -> Self {
        self.border = Some(border);
        self
    }

    pub fn with_shadow(mut self, shadow: ImageShadowEffect) -> Self {
        self.shadow = Some(shadow);
        self
    }

    pub fn with_effect_extent(mut self, extent: ImageEffectExtent) -> Self {
        self.effect_extent = Some(extent);
        self
    }

    pub fn with_alignment(mut self, alignment: &str) -> Self {
        self.alignment = Some(alignment.to_string());
        self
    }

    /// Helper to create from dimensions in inches
    pub fn from_inches(rel_id: &str, width_inches: f64, height_inches: f64) -> Self {
        const EMU_PER_INCH: i64 = 914400;
        Self::new(
            rel_id,
            (width_inches * EMU_PER_INCH as f64) as i64,
            (height_inches * EMU_PER_INCH as f64) as i64,
        )
    }
}

/// Document element (paragraph, table, or image)
#[derive(Debug, Clone)]
pub enum DocElement {
    Paragraph(Box<Paragraph>),
    Table(Table),
    Image(ImageElement),
    RawXml(String), // Raw XML content (e.g. from cover template)
}

/// Table width type
#[derive(Debug, Clone, Copy, Default)]
pub enum TableWidth {
    #[default]
    Auto,
    Dxa(u32), // Absolute width in twips
    Pct(u32), // Percentage in 50ths of a percent (5000 = 100%)
}

/// Table structure
#[derive(Debug, Clone)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub column_widths: Vec<u32>, // In twips (20ths of a point)
    pub has_header_row: bool,
    pub width: TableWidth,
    pub borders: Option<BorderStyles>, // Template borders
    pub cell_margins: Option<CellMargins>,
}

impl Table {
    pub fn new() -> Self {
        Self {
            rows: Vec::new(),
            column_widths: Vec::new(),
            has_header_row: false,
            width: TableWidth::Auto,
            borders: None,
            cell_margins: None,
        }
    }

    /// Set table borders from template
    pub fn with_borders(mut self, borders: BorderStyles) -> Self {
        self.borders = Some(borders);
        self
    }

    /// Set table cell margins from template
    pub fn with_cell_margins(mut self, margins: CellMargins) -> Self {
        self.cell_margins = Some(margins);
        self
    }

    /// Add a row to the table
    pub fn add_row(mut self, row: TableRow) -> Self {
        self.rows.push(row);
        self
    }

    /// Set column widths (in twips)
    pub fn with_column_widths(mut self, widths: Vec<u32>) -> Self {
        self.column_widths = widths;
        self
    }

    /// Set whether the table has a header row
    pub fn with_header_row(mut self, has_header: bool) -> Self {
        self.has_header_row = has_header;
        self
    }

    /// Set table width
    pub fn width(mut self, width: TableWidth) -> Self {
        self.width = width;
        self
    }
}

/// Table row
#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCellElement>,
    pub is_header: bool,
}

/// Table cell
#[derive(Debug, Clone)]
pub struct TableCellElement {
    pub paragraphs: Vec<Paragraph>,
    pub width: TableWidth,
    pub alignment: Option<String>,          // "left", "center", "right"
    pub vertical_alignment: Option<String>, // "top", "center", "bottom"
    pub shading: Option<String>,            // Fill color (hex without #)
}

impl TableRow {
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            is_header: false,
        }
    }

    /// Add a cell to the row
    pub fn add_cell(mut self, cell: TableCellElement) -> Self {
        self.cells.push(cell);
        self
    }

    /// Mark this row as a header row
    pub fn header(mut self) -> Self {
        self.is_header = true;
        self
    }
}

impl TableCellElement {
    pub fn new() -> Self {
        Self {
            paragraphs: Vec::new(),
            width: TableWidth::Auto,
            alignment: None,
            vertical_alignment: None,
            shading: None,
        }
    }

    /// Add a paragraph to the cell
    pub fn add_paragraph(mut self, p: Paragraph) -> Self {
        self.paragraphs.push(p);
        self
    }

    /// Set cell width
    pub fn width(mut self, width: TableWidth) -> Self {
        self.width = width;
        self
    }

    /// Set cell alignment
    pub fn alignment(mut self, align: &str) -> Self {
        self.alignment = Some(align.to_string());
        self
    }

    /// Set cell vertical alignment
    pub fn vertical_alignment(mut self, align: &str) -> Self {
        self.vertical_alignment = Some(align.to_string());
        self
    }

    /// Set cell shading color (hex without #)
    pub fn shading(mut self, color: &str) -> Self {
        self.shading = Some(color.to_string());
        self
    }
}

impl Default for Table {
    fn default() -> Self {
        Self::new()
    }
}

impl Default for TableRow {
    fn default() -> Self {
        Self::new()
    }
}

impl Default for TableCellElement {
    fn default() -> Self {
        Self::new()
    }
}

/// Main Document XML generator
#[derive(Debug)]
pub struct DocumentXml {
    pub elements: Vec<DocElement>,
    // Page settings (A4 default)
    pub width: u32,  // Twips (11906 for A4)
    pub height: u32, // Twips (16838 for A4)
    pub margin_top: u32,
    pub margin_right: u32,
    pub margin_bottom: u32,
    pub margin_left: u32,
    pub margin_header: u32,
    pub margin_footer: u32,
    pub header_footer_refs: HeaderFooterRefs, // Header/footer references
    pub empty_header_id: Option<String>,      // ID for empty header
    pub empty_footer_id: Option<String>,      // ID for empty footer
    pub page_num_start: Option<u32>,          // Page number start for the final section
}

impl Default for DocumentXml {
    fn default() -> Self {
        Self::new()
    }
}

impl DocumentXml {
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            width: 11906,        // A4 width in twips
            height: 16838,       // A4 height in twips
            margin_top: 1440,    // 1 inch
            margin_right: 1440,  // 1 inch
            margin_bottom: 1440, // 1 inch
            margin_left: 1440,   // 1 inch
            margin_header: 708,  // 0.5 inch
            margin_footer: 708,  // 0.5 inch
            header_footer_refs: HeaderFooterRefs::default(),
            empty_header_id: None,
            empty_footer_id: None,
            page_num_start: None,
        }
    }

    /// Add a paragraph to the document
    pub fn add_paragraph(&mut self, p: Paragraph) {
        self.elements.push(DocElement::Paragraph(Box::new(p)));
    }

    /// Add a table to the document
    pub fn add_table(&mut self, table: Table) {
        self.elements.push(DocElement::Table(table));
    }

    /// Add an image element
    pub fn add_image(&mut self, image: ImageElement) {
        self.elements.push(DocElement::Image(image));
    }

    /// Add a document element (paragraph, table, or image)
    pub fn add_element(&mut self, element: DocElement) {
        self.elements.push(element);
    }

    /// Set page size (in twips)
    pub fn page_size(mut self, width: u32, height: u32) -> Self {
        self.width = width;
        self.height = height;
        self
    }

    /// Set margins (in twips)
    pub fn margins(
        mut self,
        top: u32,
        right: u32,
        bottom: u32,
        left: u32,
        header: u32,
        footer: u32,
    ) -> Self {
        self.margin_top = top;
        self.margin_right = right;
        self.margin_bottom = bottom;
        self.margin_left = left;
        self.margin_header = header;
        self.margin_footer = footer;
        self
    }

    /// Set header/footer references
    pub fn with_header_footer(mut self, refs: HeaderFooterRefs) -> Self {
        self.header_footer_refs = refs;
        self
    }

    /// Generate XML content for word/document.xml
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

        // XML declaration
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with all Word namespaces (including 2016+ extensions to prevent compatibility mode)
        let mut doc = BytesStart::new("w:document");
        doc.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        doc.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        doc.push_attribute((
            "xmlns:wp",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ));
        doc.push_attribute((
            "xmlns:wp14",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ));
        doc.push_attribute((
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ));
        doc.push_attribute((
            "xmlns:pic",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        ));
        doc.push_attribute((
            "xmlns:mc",
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ));
        doc.push_attribute((
            "xmlns:w14",
            "http://schemas.microsoft.com/office/word/2010/wordml",
        ));
        doc.push_attribute((
            "xmlns:w15",
            "http://schemas.microsoft.com/office/word/2012/wordml",
        ));
        doc.push_attribute((
            "xmlns:w16",
            "http://schemas.microsoft.com/office/word/2018/wordml",
        ));
        doc.push_attribute((
            "xmlns:w16cex",
            "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        ));
        doc.push_attribute((
            "xmlns:w16cid",
            "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        ));
        doc.push_attribute((
            "xmlns:w16se",
            "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        ));
        doc.push_attribute((
            "xmlns:wpc",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        ));
        doc.push_attribute((
            "xmlns:wpg",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        ));
        doc.push_attribute((
            "xmlns:wpi",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        ));
        doc.push_attribute((
            "xmlns:wne",
            "http://schemas.microsoft.com/office/word/2006/wordml",
        ));
        doc.push_attribute((
            "xmlns:wps",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ));
        doc.push_attribute(("xmlns:o", "urn:schemas-microsoft-com:office:office"));
        doc.push_attribute(("xmlns:v", "urn:schemas-microsoft-com:vml"));
        doc.push_attribute(("xmlns:w10", "urn:schemas-microsoft-com:office:word"));
        doc.push_attribute(("mc:Ignorable", "w14 w15 w16se w16cid w16 w16cex wp14"));

        writer.write_event(Event::Start(doc))?;
        writer.write_event(Event::Start(BytesStart::new("w:body")))?;

        // Write elements (paragraphs, tables, images, and raw XML)
        for element in &self.elements {
            match element {
                DocElement::Paragraph(p) => self.write_paragraph(&mut writer, p)?,
                DocElement::Table(table) => self.write_table(&mut writer, table)?,
                DocElement::Image(image) => {
                    // Images need to be wrapped in a paragraph and run
                    writer.write_event(Event::Start(BytesStart::new("w:p")))?;

                    // Add pPr with spacing and alignment
                    writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

                    // Alignment (w:jc)
                    if let Some(ref align) = image.alignment {
                        let mut jc = BytesStart::new("w:jc");
                        jc.push_attribute(("w:val", align.as_str()));
                        writer.write_event(Event::Empty(jc))?;
                    }

                    // Spacing
                    let mut spacing = BytesStart::new("w:spacing");
                    spacing.push_attribute(("w:before", "0"));
                    spacing.push_attribute(("w:after", "0"));
                    spacing.push_attribute(("w:line", "240"));
                    spacing.push_attribute(("w:lineRule", "auto"));
                    writer.write_event(Event::Empty(spacing))?;

                    writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;

                    writer.write_event(Event::Start(BytesStart::new("w:r")))?;
                    self.write_drawing(&mut writer, image)?;
                    writer.write_event(Event::End(BytesEnd::new("w:r")))?;
                    writer.write_event(Event::End(BytesEnd::new("w:p")))?;
                }
                DocElement::RawXml(xml) => {
                    self.write_raw_xml(&mut writer, xml)?;
                }
            }
        }

        // Section properties (Page size/margins)
        self.write_sect_pr(&mut writer)?;

        writer.write_event(Event::End(BytesEnd::new("w:body")))?;
        writer.write_event(Event::End(BytesEnd::new("w:document")))?;

        Ok(writer.into_inner().into_inner())
    }

    /// Write raw XML content (e.g. from cover template)
    fn write_raw_xml<W: std::io::Write>(&self, writer: &mut Writer<W>, xml: &str) -> Result<()> {
        // Wrap the XML in a wrapper to make it valid XML for parsing
        let wrapped = format!("<wrapper>{}</wrapper>", xml);

        // Parse the wrapped XML
        let mut reader = quick_xml::Reader::from_reader(wrapped.as_bytes());
        reader.config_mut().trim_text_end = false;

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = std::str::from_utf8(e.name().into_inner()).unwrap_or("");
                    if name != "wrapper" {
                        writer.write_event(Event::Start(e.to_owned()))?;
                    }
                }
                Ok(Event::End(e)) => {
                    let name = std::str::from_utf8(e.name().into_inner()).unwrap_or("");
                    if name != "wrapper" {
                        writer.write_event(Event::End(e.to_owned()))?;
                    }
                }
                Ok(Event::Empty(e)) => {
                    let name = std::str::from_utf8(e.name().into_inner()).unwrap_or("");
                    if name != "wrapper" {
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
                }
                Ok(Event::Text(e)) => {
                    writer.write_event(Event::Text(e.to_owned()))?;
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(())
    }

    /// Write a paragraph element
    fn write_paragraph<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        p: &Paragraph,
    ) -> Result<()> {
        // Case 1: Section break with header/footer suppression
        // For cover/TOC sections, we want NO headers/footers at all.
        // If empty header/footer IDs are set, use them to explicitly prevent inheritance.
        // Otherwise, passing None means no w:headerReference/w:footerReference elements,
        // which tells Word this section has no headers/footers.
        if p.section_break.is_some() && p.suppress_header_footer {
            if self.empty_header_id.is_some() || self.empty_footer_id.is_some() {
                // Create refs with empty header/footer IDs to explicitly prevent inheritance
                let empty_refs = HeaderFooterRefs {
                    default_header_id: self.empty_header_id.clone(),
                    default_footer_id: self.empty_footer_id.clone(),
                    first_header_id: None,
                    first_footer_id: None,
                    different_first_page: false,
                };
                return p.write_xml(writer, Some(&empty_refs));
            } else {
                return p.write_xml(writer, None);
            }
        }

        // Case 2: Normal section break (inherit or use current refs)
        let refs = if p.section_break.is_some() {
            Some(&self.header_footer_refs)
        } else {
            None
        };
        p.write_xml(writer, refs)
    }

    /// Write a run element
    #[allow(dead_code)]
    fn write_run<W: std::io::Write>(&self, writer: &mut Writer<W>, run: &Run) -> Result<()> {
        run.write_xml(writer)
    }

    /// Write section properties (page size and margins)
    fn write_sect_pr<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:sectPr")))?;

        // Default header
        if let Some(ref id) = self.header_footer_refs.default_header_id {
            let mut header_ref = BytesStart::new("w:headerReference");
            header_ref.push_attribute(("w:type", "default"));
            header_ref.push_attribute(("r:id", id.as_str()));
            writer.write_event(Event::Empty(header_ref))?;
        }

        // First page header (if different first page enabled)
        if self.header_footer_refs.different_first_page {
            if let Some(ref id) = self.header_footer_refs.first_header_id {
                let mut header_ref = BytesStart::new("w:headerReference");
                header_ref.push_attribute(("w:type", "first"));
                header_ref.push_attribute(("r:id", id.as_str()));
                writer.write_event(Event::Empty(header_ref))?;
            }
        }

        // Default footer
        if let Some(ref id) = self.header_footer_refs.default_footer_id {
            let mut footer_ref = BytesStart::new("w:footerReference");
            footer_ref.push_attribute(("w:type", "default"));
            footer_ref.push_attribute(("r:id", id.as_str()));
            writer.write_event(Event::Empty(footer_ref))?;
        }

        // First page footer (if different first page enabled)
        if self.header_footer_refs.different_first_page {
            if let Some(ref id) = self.header_footer_refs.first_footer_id {
                let mut footer_ref = BytesStart::new("w:footerReference");
                footer_ref.push_attribute(("w:type", "first"));
                footer_ref.push_attribute(("r:id", id.as_str()));
                writer.write_event(Event::Empty(footer_ref))?;
            }
        }

        // Page numbering (restart at specific number if set)
        if let Some(start) = self.page_num_start {
            let mut pg_num = BytesStart::new("w:pgNumType");
            pg_num.push_attribute(("w:start", start.to_string().as_str()));
            writer.write_event(Event::Empty(pg_num))?;
        }

        // Page size
        let mut pg_sz = BytesStart::new("w:pgSz");
        pg_sz.push_attribute(("w:w", self.width.to_string().as_str()));
        pg_sz.push_attribute(("w:h", self.height.to_string().as_str()));
        writer.write_event(Event::Empty(pg_sz))?;

        // Margins
        let mut pg_mar = BytesStart::new("w:pgMar");
        pg_mar.push_attribute(("w:top", self.margin_top.to_string().as_str()));
        pg_mar.push_attribute(("w:right", self.margin_right.to_string().as_str()));
        pg_mar.push_attribute(("w:bottom", self.margin_bottom.to_string().as_str()));
        pg_mar.push_attribute(("w:left", self.margin_left.to_string().as_str()));
        pg_mar.push_attribute(("w:header", self.margin_header.to_string().as_str()));
        pg_mar.push_attribute(("w:footer", self.margin_footer.to_string().as_str()));
        pg_mar.push_attribute(("w:gutter", "0"));
        writer.write_event(Event::Empty(pg_mar))?;

        // Columns (single column by default)
        let mut cols = BytesStart::new("w:cols");
        cols.push_attribute(("w:space", "708")); // 0.5 inch
        writer.write_event(Event::Empty(cols))?;

        // Title page (different first page)
        if self.header_footer_refs.different_first_page {
            writer.write_event(Event::Empty(BytesStart::new("w:titlePg")))?;
        }

        // Document grid (for Asian text)
        let mut doc_grid = BytesStart::new("w:docGrid");
        doc_grid.push_attribute(("w:linePitch", "360"));
        writer.write_event(Event::Empty(doc_grid))?;

        writer.write_event(Event::End(BytesEnd::new("w:sectPr")))?;
        Ok(())
    }

    pub fn write_border<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        tag: &str,
        border: &BorderStyle,
    ) -> Result<()> {
        let mut elem = BytesStart::new(tag);
        elem.push_attribute(("w:val", border.style.as_str()));
        elem.push_attribute(("w:sz", border.width.to_string().as_str()));
        elem.push_attribute(("w:space", "0"));
        // Remove # from color if present
        let color = border.color.trim_start_matches('#');
        elem.push_attribute(("w:color", color));
        writer.write_event(Event::Empty(elem))?;
        Ok(())
    }

    /// Write a table element
    fn write_table<W: std::io::Write>(&self, writer: &mut Writer<W>, table: &Table) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tbl")))?;

        // Table properties
        writer.write_event(Event::Start(BytesStart::new("w:tblPr")))?;

        // Table style
        let mut tbl_style = BytesStart::new("w:tblStyle");
        tbl_style.push_attribute(("w:val", "TableGrid"));
        writer.write_event(Event::Empty(tbl_style))?;

        // Table width
        let mut tbl_w = BytesStart::new("w:tblW");
        match table.width {
            TableWidth::Auto => {
                tbl_w.push_attribute(("w:w", "0"));
                tbl_w.push_attribute(("w:type", "auto"));
            }
            TableWidth::Dxa(w) => {
                tbl_w.push_attribute(("w:w", w.to_string().as_str()));
                tbl_w.push_attribute(("w:type", "dxa"));
            }
            TableWidth::Pct(w) => {
                tbl_w.push_attribute(("w:w", w.to_string().as_str()));
                tbl_w.push_attribute(("w:type", "pct"));
            }
        }
        writer.write_event(Event::Empty(tbl_w))?;

        // Table cell margins (padding)
        writer.write_event(Event::Start(BytesStart::new("w:tblCellMar")))?;

        let (top, bottom, left, right) = if let Some(margins) = &table.cell_margins {
            (margins.top, margins.bottom, margins.left, margins.right)
        } else {
            (100, 100, 100, 100) // Default values
        };

        // Top margin
        let mut top_mar = BytesStart::new("w:top");
        top_mar.push_attribute(("w:w", top.to_string().as_str()));
        top_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(top_mar))?;

        // Bottom margin
        let mut bottom_mar = BytesStart::new("w:bottom");
        bottom_mar.push_attribute(("w:w", bottom.to_string().as_str()));
        bottom_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(bottom_mar))?;

        // Left margin
        let mut left_mar = BytesStart::new("w:left");
        left_mar.push_attribute(("w:w", left.to_string().as_str()));
        left_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(left_mar))?;

        // Right margin
        let mut right_mar = BytesStart::new("w:right");
        right_mar.push_attribute(("w:w", right.to_string().as_str()));
        right_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(right_mar))?;

        writer.write_event(Event::End(BytesEnd::new("w:tblCellMar")))?;

        // Table borders
        writer.write_event(Event::Start(BytesStart::new("w:tblBorders")))?;

        if let Some(borders) = &table.borders {
            // Use template borders
            self.write_border(writer, "w:top", &borders.top)?;
            self.write_border(writer, "w:left", &borders.left)?;
            self.write_border(writer, "w:bottom", &borders.bottom)?;
            self.write_border(writer, "w:right", &borders.right)?;
            self.write_border(writer, "w:insideH", &borders.inside_h)?;
            self.write_border(writer, "w:insideV", &borders.inside_v)?;
        } else {
            // Default borders
            // Top border
            let mut border = BytesStart::new("w:top");
            border.push_attribute(("w:val", "single"));
            border.push_attribute(("w:sz", "4"));
            border.push_attribute(("w:space", "0"));
            border.push_attribute(("w:color", "auto"));
            writer.write_event(Event::Empty(border))?;

            // Left border
            let mut border = BytesStart::new("w:left");
            border.push_attribute(("w:val", "single"));
            border.push_attribute(("w:sz", "4"));
            border.push_attribute(("w:space", "0"));
            border.push_attribute(("w:color", "auto"));
            writer.write_event(Event::Empty(border))?;

            // Bottom border
            let mut border = BytesStart::new("w:bottom");
            border.push_attribute(("w:val", "single"));
            border.push_attribute(("w:sz", "4"));
            border.push_attribute(("w:space", "0"));
            border.push_attribute(("w:color", "auto"));
            writer.write_event(Event::Empty(border))?;

            // Right border
            let mut border = BytesStart::new("w:right");
            border.push_attribute(("w:val", "single"));
            border.push_attribute(("w:sz", "4"));
            border.push_attribute(("w:space", "0"));
            border.push_attribute(("w:color", "auto"));
            writer.write_event(Event::Empty(border))?;

            // Inside horizontal borders
            let mut border = BytesStart::new("w:insideH");
            border.push_attribute(("w:val", "single"));
            border.push_attribute(("w:sz", "4"));
            border.push_attribute(("w:space", "0"));
            border.push_attribute(("w:color", "auto"));
            writer.write_event(Event::Empty(border))?;

            // Inside vertical borders
            let mut border = BytesStart::new("w:insideV");
            border.push_attribute(("w:val", "single"));
            border.push_attribute(("w:sz", "4"));
            border.push_attribute(("w:space", "0"));
            border.push_attribute(("w:color", "auto"));
            writer.write_event(Event::Empty(border))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tblBorders")))?;
        writer.write_event(Event::End(BytesEnd::new("w:tblPr")))?;

        // Table grid (column widths)
        writer.write_event(Event::Start(BytesStart::new("w:tblGrid")))?;
        for width in &table.column_widths {
            let mut grid_col = BytesStart::new("w:gridCol");
            grid_col.push_attribute(("w:w", width.to_string().as_str()));
            writer.write_event(Event::Empty(grid_col))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:tblGrid")))?;

        // Write rows
        for row in &table.rows {
            self.write_table_row(writer, row)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tbl")))?;
        Ok(())
    }

    /// Write a table row element
    fn write_table_row<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        row: &TableRow,
    ) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tr")))?;

        // Row properties (optional)
        if row.is_header {
            writer.write_event(Event::Start(BytesStart::new("w:trPr")))?;
            writer.write_event(Event::Empty(BytesStart::new("w:tblHeader")))?;
            writer.write_event(Event::End(BytesEnd::new("w:trPr")))?;
        }

        // Write cells
        for cell in &row.cells {
            self.write_table_cell(writer, cell)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tr")))?;
        Ok(())
    }

    /// Write a table cell element
    fn write_table_cell<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        cell: &TableCellElement,
    ) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:tc")))?;

        // Cell properties
        writer.write_event(Event::Start(BytesStart::new("w:tcPr")))?;

        // Cell width
        let mut tc_w = BytesStart::new("w:tcW");
        match cell.width {
            TableWidth::Auto => {
                tc_w.push_attribute(("w:w", "0"));
                tc_w.push_attribute(("w:type", "auto"));
            }
            TableWidth::Dxa(w) => {
                tc_w.push_attribute(("w:w", w.to_string().as_str()));
                tc_w.push_attribute(("w:type", "dxa"));
            }
            TableWidth::Pct(w) => {
                tc_w.push_attribute(("w:w", w.to_string().as_str()));
                tc_w.push_attribute(("w:type", "pct"));
            }
        }
        writer.write_event(Event::Empty(tc_w))?;

        // Cell alignment
        if let Some(align) = &cell.alignment {
            let mut jc = BytesStart::new("w:jc");
            jc.push_attribute(("w:val", align.as_str()));
            writer.write_event(Event::Empty(jc))?;
        }

        // Cell vertical alignment
        if let Some(v_align) = &cell.vertical_alignment {
            let mut valign = BytesStart::new("w:vAlign");
            valign.push_attribute(("w:val", v_align.as_str()));
            writer.write_event(Event::Empty(valign))?;
        }

        // Cell shading
        if let Some(shading) = &cell.shading {
            let mut shd = BytesStart::new("w:shd");
            shd.push_attribute(("w:val", "clear"));
            shd.push_attribute(("w:color", "auto"));
            shd.push_attribute(("w:fill", shading.as_str()));
            writer.write_event(Event::Empty(shd))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tcPr")))?;

        // Write paragraphs in the cell
        for p in &cell.paragraphs {
            p.write_xml(writer, None)?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:tc")))?;
        Ok(())
    }

    /// Write a drawing element (for images)
    fn write_drawing<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        image: &ImageElement,
    ) -> Result<()> {
        // <w:drawing>
        writer.write_event(Event::Start(BytesStart::new("w:drawing")))?;

        // <wp:inline distT="0" distB="0" distL="0" distR="0">
        let mut inline = BytesStart::new("wp:inline");
        inline.push_attribute(("distT", "0"));
        inline.push_attribute(("distB", "0"));
        inline.push_attribute(("distL", "0"));
        inline.push_attribute(("distR", "0"));
        writer.write_event(Event::Start(inline))?;

        // <wp:extent cx="WIDTH" cy="HEIGHT"/>
        let mut extent = BytesStart::new("wp:extent");
        extent.push_attribute(("cx", image.width_emu.to_string().as_str()));
        extent.push_attribute(("cy", image.height_emu.to_string().as_str()));
        writer.write_event(Event::Empty(extent))?;

        // <wp:effectExtent l="0" t="0" r="0" b="0"/>
        let extent = image.effect_extent.as_ref();
        let mut effect = BytesStart::new("wp:effectExtent");
        effect.push_attribute((
            "l",
            extent
                .map_or("0".to_string(), |e| e.left.to_string())
                .as_str(),
        ));
        effect.push_attribute((
            "t",
            extent
                .map_or("0".to_string(), |e| e.top.to_string())
                .as_str(),
        ));
        effect.push_attribute((
            "r",
            extent
                .map_or("0".to_string(), |e| e.right.to_string())
                .as_str(),
        ));
        effect.push_attribute((
            "b",
            extent
                .map_or("0".to_string(), |e| e.bottom.to_string())
                .as_str(),
        ));
        writer.write_event(Event::Empty(effect))?;

        // <wp:docPr id="1" name="Picture 1" descr="alt text"/>
        let mut doc_pr = BytesStart::new("wp:docPr");
        doc_pr.push_attribute(("id", image.id.to_string().as_str()));
        doc_pr.push_attribute(("name", format!("Picture {}", image.id).as_str()));
        if !image.alt_text.is_empty() {
            doc_pr.push_attribute(("descr", image.alt_text.as_str()));
        }
        writer.write_event(Event::Empty(doc_pr))?;

        // <wp:cNvGraphicFramePr>
        writer.write_event(Event::Start(BytesStart::new("wp:cNvGraphicFramePr")))?;
        // <a:graphicFrameLocks noChangeAspect="1"/>
        let mut locks = BytesStart::new("a:graphicFrameLocks");
        locks.push_attribute((
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ));
        locks.push_attribute(("noChangeAspect", "1"));
        writer.write_event(Event::Empty(locks))?;
        writer.write_event(Event::End(BytesEnd::new("wp:cNvGraphicFramePr")))?;

        // <a:graphic>
        writer.write_event(Event::Start(BytesStart::new("a:graphic")))?;
        // <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        let mut data = BytesStart::new("a:graphicData");
        data.push_attribute((
            "uri",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        ));
        writer.write_event(Event::Start(data))?;

        // <pic:pic>
        writer.write_event(Event::Start(BytesStart::new("pic:pic")))?;

        // <pic:nvPicPr>
        writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;
        // <pic:cNvPr id="0" name="Picture 0"/>
        let mut c_nv_pr = BytesStart::new("pic:cNvPr");
        c_nv_pr.push_attribute(("id", image.id.to_string().as_str()));
        c_nv_pr.push_attribute(("name", format!("Picture {}", image.id).as_str()));
        writer.write_event(Event::Empty(c_nv_pr))?;
        // <pic:cNvPicPr/>
        writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

        // <pic:blipFill>
        writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
        // <a:blip r:embed="rId4"/>
        let mut blip = BytesStart::new("a:blip");
        blip.push_attribute(("r:embed", image.rel_id.as_str()));
        writer.write_event(Event::Empty(blip))?;
        // <a:stretch><a:fillRect/></a:stretch>
        writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
        writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
        writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

        // <pic:spPr>
        writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;
        // <a:xfrm><a:off x="0" y="0"/><a:ext cx="WIDTH" cy="HEIGHT"/></a:xfrm>
        writer.write_event(Event::Start(BytesStart::new("a:xfrm")))?;
        let mut off = BytesStart::new("a:off");
        off.push_attribute(("x", "0"));
        off.push_attribute(("y", "0"));
        writer.write_event(Event::Empty(off))?;
        let mut ext = BytesStart::new("a:ext");
        ext.push_attribute(("cx", image.width_emu.to_string().as_str()));
        ext.push_attribute(("cy", image.height_emu.to_string().as_str()));
        writer.write_event(Event::Empty(ext))?;
        writer.write_event(Event::End(BytesEnd::new("a:xfrm")))?;
        // <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        let mut geom = BytesStart::new("a:prstGeom");
        geom.push_attribute(("prst", "rect"));
        writer.write_event(Event::Start(geom))?;
        writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;

        // <a:ln> (border)
        if let Some(border) = &image.border {
            let mut ln = BytesStart::new("a:ln");
            if let Some(w) = border.width {
                ln.push_attribute(("w", w.to_string().as_str()));
            }
            writer.write_event(Event::Start(ln))?;

            if border.fill_type == "solid" {
                writer.write_event(Event::Start(BytesStart::new("a:solidFill")))?;
                if border.is_scheme_color {
                    let mut clr = BytesStart::new("a:schemeClr");
                    clr.push_attribute(("val", border.color.as_str()));
                    writer.write_event(Event::Empty(clr))?;
                } else {
                    let mut clr = BytesStart::new("a:srgbClr");
                    clr.push_attribute(("val", border.color.as_str()));
                    writer.write_event(Event::Empty(clr))?;
                }
                writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
            } else if border.fill_type == "none" {
                writer.write_event(Event::Empty(BytesStart::new("a:noFill")))?;
            }

            writer.write_event(Event::End(BytesEnd::new("a:ln")))?;
        }

        // <a:effectLst> (shadow)
        if let Some(shadow) = &image.shadow {
            writer.write_event(Event::Start(BytesStart::new("a:effectLst")))?;
            let mut outer_shadow = BytesStart::new("a:outerShdw");
            outer_shadow.push_attribute(("blurRad", shadow.blur_radius.to_string().as_str()));
            outer_shadow.push_attribute(("dist", shadow.distance.to_string().as_str()));
            outer_shadow.push_attribute(("dir", shadow.direction.to_string().as_str()));
            outer_shadow.push_attribute(("algn", shadow.alignment.as_str()));
            writer.write_event(Event::Start(outer_shadow))?;

            let mut clr = BytesStart::new("a:srgbClr");
            clr.push_attribute(("val", shadow.color.as_str()));
            writer.write_event(Event::Start(clr))?;

            let mut alpha = BytesStart::new("a:alpha");
            alpha.push_attribute(("val", shadow.alpha.to_string().as_str()));
            writer.write_event(Event::Empty(alpha))?;

            writer.write_event(Event::End(BytesEnd::new("a:srgbClr")))?;
            writer.write_event(Event::End(BytesEnd::new("a:outerShdw")))?;
            writer.write_event(Event::End(BytesEnd::new("a:effectLst")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

        writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
        writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;
        writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_paragraph_to_xml() {
        let mut p = Paragraph::with_style("Heading1");
        p = p.add_text("Hello World");
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        p.write_xml(&mut writer, None).unwrap();
        let xml = String::from_utf8(writer.into_inner().into_inner()).unwrap();
        assert!(xml.contains("<w:pStyle w:val=\"Heading1\"/>"));
        assert!(xml.contains("<w:t xml:space=\"preserve\">Hello World</w:t>"));
    }

    #[test]
    fn test_table_to_xml() {
        let table =
            Table::new()
                .with_header_row(true)
                .add_row(TableRow::new().header().add_cell(
                    TableCellElement::new().add_paragraph(Paragraph::new().add_text("A1")),
                ))
                .add_row(TableRow::new().add_cell(
                    TableCellElement::new().add_paragraph(Paragraph::new().add_text("A2")),
                ));

        let mut doc = DocumentXml::new();
        doc.add_table(table);
        let xml = String::from_utf8(doc.to_xml().unwrap()).unwrap();
        assert!(xml.contains("<w:tblHeader/>"));
        assert!(xml.contains("A1"));
        assert!(xml.contains("A2"));
    }

    #[test]
    fn test_image_to_xml() {
        let image = ImageElement::new("rId1", 1000, 1000);
        let mut doc = DocumentXml::new();
        doc.add_image(image);
        let xml = String::from_utf8(doc.to_xml().unwrap()).unwrap();
        assert!(xml.contains("<w:drawing>"));
        assert!(xml.contains("r:embed=\"rId1\""));
    }

    #[test]
    fn test_image_with_effects() {
        let image = ImageElement::new("rId1", 1000000, 750000)
            .with_border(ImageBorderEffect {
                fill_type: "solid".to_string(),
                color: "accent1".to_string(),
                is_scheme_color: true,
                width: None,
            })
            .with_shadow(ImageShadowEffect {
                blur_radius: 190500,
                distance: 228600,
                direction: 2700000,
                alignment: "ctr".to_string(),
                color: "000000".to_string(),
                alpha: 30000,
            })
            .with_effect_extent(ImageEffectExtent {
                left: 38100,
                top: 38100,
                right: 326390,
                bottom: 327660,
            });

        let mut doc = DocumentXml::new();
        doc.add_image(image);
        let xml = String::from_utf8(doc.to_xml().unwrap()).unwrap();

        assert!(xml.contains("<a:ln>"));
        assert!(xml.contains("<a:schemeClr"));
        assert!(xml.contains("<a:outerShdw"));
        assert!(xml.contains("blurRad=\"190500\""));
        assert!(xml.contains("<a:alpha val=\"30000\""));
    }
}
