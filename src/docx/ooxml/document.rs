//! Generate word/document.xml for DOCX

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;
use crate::i18n::detection::detect_language;

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

                // Page size
                let mut pg_sz = BytesStart::new("w:pgSz");
                pg_sz.push_attribute(("w:w", "11906"));
                pg_sz.push_attribute(("w:h", "16838"));
                writer.write_event(Event::Empty(pg_sz))?;

                // Margins
                let mut pg_mar = BytesStart::new("w:pgMar");
                pg_mar.push_attribute(("w:top", "1440"));
                pg_mar.push_attribute(("w:right", "1440"));
                pg_mar.push_attribute(("w:bottom", "1440"));
                pg_mar.push_attribute(("w:left", "1440"));
                pg_mar.push_attribute(("w:header", "708"));
                pg_mar.push_attribute(("w:footer", "708"));
                pg_mar.push_attribute(("w:gutter", "0"));
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
    pub alignment: Option<String>, // "left", "center", "right"
    pub shading: Option<String>,   // Fill color (hex without #)
}

impl Table {
    pub fn new() -> Self {
        Self {
            rows: Vec::new(),
            column_widths: Vec::new(),
            has_header_row: false,
            width: TableWidth::Auto,
        }
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

        // Write elements (paragraphs, tables, and images)
        for element in &self.elements {
            match element {
                DocElement::Paragraph(p) => self.write_paragraph(&mut writer, p)?,
                DocElement::Table(table) => self.write_table(&mut writer, table)?,
                DocElement::Image(image) => {
                    // Images need to be wrapped in a paragraph and run
                    writer.write_event(Event::Start(BytesStart::new("w:p")))?;

                    // Add pPr with spacing 0 and single line spacing
                    writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;
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
            }
        }

        // Section properties (Page size/margins)
        self.write_sect_pr(&mut writer)?;

        writer.write_event(Event::End(BytesEnd::new("w:body")))?;
        writer.write_event(Event::End(BytesEnd::new("w:document")))?;

        Ok(writer.into_inner().into_inner())
    }

    /// Write a paragraph element
    fn write_paragraph<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        p: &Paragraph,
    ) -> Result<()> {
        // Case 1: Section break with header/footer suppression
        if p.section_break.is_some() && p.suppress_header_footer {
            if self.empty_header_id.is_some() || self.empty_footer_id.is_some() {
                // Create temporary refs pointing to empty files
                let mut empty_refs = HeaderFooterRefs::default();
                if let Some(id) = &self.empty_header_id {
                    empty_refs.default_header_id = Some(id.clone());
                }
                if let Some(id) = &self.empty_footer_id {
                    empty_refs.default_footer_id = Some(id.clone());
                }
                return p.write_xml(writer, Some(&empty_refs));
            } else {
                // No empty files available, just suppress by passing None
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

        // Top margin: 100 twips (5pt)
        let mut top_mar = BytesStart::new("w:top");
        top_mar.push_attribute(("w:w", "100"));
        top_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(top_mar))?;

        // Bottom margin: 100 twips (5pt)
        let mut bottom_mar = BytesStart::new("w:bottom");
        bottom_mar.push_attribute(("w:w", "100"));
        bottom_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(bottom_mar))?;

        // Left margin: 100 twips (5pt)
        let mut left_mar = BytesStart::new("w:left");
        left_mar.push_attribute(("w:w", "100"));
        left_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(left_mar))?;

        // Right margin: 100 twips (5pt)
        let mut right_mar = BytesStart::new("w:right");
        right_mar.push_attribute(("w:w", "100"));
        right_mar.push_attribute(("w:type", "dxa"));
        writer.write_event(Event::Empty(right_mar))?;

        writer.write_event(Event::End(BytesEnd::new("w:tblCellMar")))?;

        // Table borders
        writer.write_event(Event::Start(BytesStart::new("w:tblBorders")))?;

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
        let mut effect = BytesStart::new("wp:effectExtent");
        effect.push_attribute(("l", "0"));
        effect.push_attribute(("t", "0"));
        effect.push_attribute(("r", "0"));
        effect.push_attribute(("b", "0"));
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
        let mut locks = BytesStart::new("a:graphicFrameLocks");
        locks.push_attribute((
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ));
        locks.push_attribute(("noChangeAspect", "1"));
        writer.write_event(Event::Empty(locks))?;
        writer.write_event(Event::End(BytesEnd::new("wp:cNvGraphicFramePr")))?;

        // <a:graphic>
        let mut graphic = BytesStart::new("a:graphic");
        graphic.push_attribute((
            "xmlns:a",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
        ));
        writer.write_event(Event::Start(graphic))?;

        // <a:graphicData uri="...">
        let mut graphic_data = BytesStart::new("a:graphicData");
        graphic_data.push_attribute((
            "uri",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        ));
        writer.write_event(Event::Start(graphic_data))?;

        // <pic:pic>
        let mut pic_pic = BytesStart::new("pic:pic");
        pic_pic.push_attribute((
            "xmlns:pic",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        ));
        writer.write_event(Event::Start(pic_pic))?;

        // <pic:nvPicPr>
        writer.write_event(Event::Start(BytesStart::new("pic:nvPicPr")))?;
        let mut cnv_pr = BytesStart::new("pic:cNvPr");
        cnv_pr.push_attribute(("id", image.id.to_string().as_str()));
        cnv_pr.push_attribute(("name", image.name.as_str()));
        writer.write_event(Event::Empty(cnv_pr))?;
        writer.write_event(Event::Empty(BytesStart::new("pic:cNvPicPr")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:nvPicPr")))?;

        // <pic:blipFill>
        writer.write_event(Event::Start(BytesStart::new("pic:blipFill")))?;
        let mut blip = BytesStart::new("a:blip");
        blip.push_attribute(("r:embed", image.rel_id.as_str()));
        writer.write_event(Event::Empty(blip))?;
        writer.write_event(Event::Start(BytesStart::new("a:stretch")))?;
        writer.write_event(Event::Empty(BytesStart::new("a:fillRect")))?;
        writer.write_event(Event::End(BytesEnd::new("a:stretch")))?;
        writer.write_event(Event::End(BytesEnd::new("pic:blipFill")))?;

        // <pic:spPr>
        writer.write_event(Event::Start(BytesStart::new("pic:spPr")))?;

        // <a:xfrm>
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

        // <a:prstGeom prst="rect">
        let mut prst_geom = BytesStart::new("a:prstGeom");
        prst_geom.push_attribute(("prst", "rect"));
        writer.write_event(Event::Start(prst_geom))?;
        writer.write_event(Event::Empty(BytesStart::new("a:avLst")))?;
        writer.write_event(Event::End(BytesEnd::new("a:prstGeom")))?;

        writer.write_event(Event::End(BytesEnd::new("pic:spPr")))?;

        // Close all tags
        writer.write_event(Event::End(BytesEnd::new("pic:pic")))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphicData")))?;
        writer.write_event(Event::End(BytesEnd::new("a:graphic")))?;
        writer.write_event(Event::End(BytesEnd::new("wp:inline")))?;
        writer.write_event(Event::End(BytesEnd::new("w:drawing")))?;

        Ok(())
    }
}

impl Default for DocumentXml {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_run_builder() {
        let run = Run::new("Hello")
            .bold()
            .italic()
            .underline()
            .size(24)
            .color("FF0000");

        assert_eq!(run.text, "Hello");
        assert!(run.bold);
        assert!(run.italic);
        assert!(run.underline);
        assert_eq!(run.size, Some(24));
        assert_eq!(run.color, Some("FF0000".to_string()));
    }

    #[test]
    fn test_paragraph_builder() {
        let p = Paragraph::with_style("Heading1")
            .add_text("Chapter 1")
            .align("center")
            .spacing(240, 60);

        assert_eq!(p.style_id, Some("Heading1".to_string()));
        assert_eq!(p.children.len(), 1);
        match &p.children[0] {
            ParagraphChild::Run(run) => {
                assert_eq!(run.text, "Chapter 1");
            }
            _ => panic!("Expected Run child"),
        }
        assert_eq!(p.align, Some("center".to_string()));
        assert_eq!(p.spacing_before, Some(240));
        assert_eq!(p.spacing_after, Some(60));
    }

    #[test]
    fn test_paragraph_with_numbering() {
        let p = Paragraph::new().add_text("First item").numbering(1, 0);

        assert_eq!(p.numbering_id, Some(1));
        assert_eq!(p.numbering_level, Some(0));
    }

    #[test]
    fn test_document_xml_basic() {
        let mut doc = DocumentXml::new();
        doc.add_paragraph(Paragraph::with_style("Normal").add_text("Hello, World!"));

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<?xml version"));
        assert!(xml_str.contains("<w:document"));
        assert!(xml_str.contains("<w:body>"));
        assert!(xml_str.contains("<w:p>"));
        assert!(xml_str.contains("<w:pStyle w:val=\"Normal\"/>"));
        assert!(xml_str.contains("Hello, World!"));
        assert!(xml_str.contains("<w:sectPr>"));
    }

    #[test]
    fn test_document_xml_multiple_paragraphs() {
        let mut doc = DocumentXml::new();
        doc.add_paragraph(Paragraph::with_style("Heading1").add_text("Title"));
        doc.add_paragraph(Paragraph::with_style("Normal").add_text("Content"));

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Count paragraph elements
        let p_count = xml_str.matches("<w:p>").count();
        assert_eq!(p_count, 2);
    }

    #[test]
    fn test_run_formatting() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::new()
            .add_run(Run::new("Normal "))
            .add_run(Run::new("Bold").bold())
            .add_run(Run::new(" Italic").italic());
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:b/>"));
        assert!(xml_str.contains("<w:i/>"));
        assert!(xml_str.contains("Normal"));
        assert!(xml_str.contains("Bold"));
        assert!(xml_str.contains("Italic"));
    }

    #[test]
    fn test_preserve_space() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::new().add_run(Run::new("  Multiple   spaces  ").preserve_space(true));
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("xml:space=\"preserve\""));
    }

    #[test]
    fn test_page_size_and_margins() {
        let doc = DocumentXml::new()
            .page_size(11906, 16838)
            .margins(1440, 1440, 1440, 1440, 708, 708);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:pgSz w:w=\"11906\" w:h=\"16838\"/>"));
        assert!(xml_str.contains("<w:pgMar"));
        assert!(xml_str.contains("w:top=\"1440\""));
        assert!(xml_str.contains("w:right=\"1440\""));
        assert!(xml_str.contains("w:bottom=\"1440\""));
        assert!(xml_str.contains("w:left=\"1440\""));
        assert!(xml_str.contains("w:header=\"708\""));
        assert!(xml_str.contains("w:footer=\"708\""));
    }

    #[test]
    fn test_paragraph_keep_with_next() {
        let p = Paragraph::new().add_text("Keep with next").keep_with_next();

        assert!(p.keep_with_next);
    }

    #[test]
    fn test_paragraph_page_break_before() {
        let p = Paragraph::new().add_text("New page").page_break_before();

        assert!(p.page_break_before);
    }

    #[test]
    fn test_run_strike() {
        let run = Run::new("Deleted").strike();
        assert!(run.strike);
    }

    #[test]
    fn test_run_highlight() {
        let run = Run::new("Highlighted").highlight("yellow");
        assert_eq!(run.highlight, Some("yellow".to_string()));
    }

    #[test]
    fn test_run_font_override() {
        let run = Run::new("Custom font").font("Arial");
        assert_eq!(run.font, Some("Arial".to_string()));
    }

    #[test]
    fn test_xml_namespaces() {
        let doc = DocumentXml::new();
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str
            .contains("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""));
        assert!(xml_str.contains(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        ));
        assert!(xml_str.contains(
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
        ));
        assert!(
            xml_str.contains("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"")
        );
        assert!(xml_str
            .contains("xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""));
    }

    #[test]
    fn test_empty_paragraph() {
        let mut doc = DocumentXml::new();
        doc.add_paragraph(Paragraph::new());

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Should still have paragraph element even if empty
        assert!(xml_str.contains("<w:p>"));
    }

    #[test]
    fn test_paragraph_indent() {
        let p = Paragraph::new().add_text("Indented text").indent(720);

        assert_eq!(p.indent_left, Some(720));
    }

    #[test]
    fn test_complex_paragraph() {
        let p = Paragraph::with_style("Heading1")
            .add_run(Run::new("Chapter ").bold())
            .add_run(Run::new("1").size(32).color("2F5496"))
            .align("center")
            .spacing(240, 60)
            .keep_with_next();

        assert_eq!(p.style_id, Some("Heading1".to_string()));
        assert_eq!(p.children.len(), 2);
        match &p.children[0] {
            ParagraphChild::Run(run) => {
                assert!(run.bold);
            }
            _ => panic!("Expected Run child"),
        }
        match &p.children[1] {
            ParagraphChild::Run(run) => {
                assert_eq!(run.size, Some(32));
                assert_eq!(run.color, Some("2F5496".to_string()));
            }
            _ => panic!("Expected Run child"),
        }
        assert_eq!(p.align, Some("center".to_string()));
        assert!(p.keep_with_next);
    }

    #[test]
    fn test_default_implementations() {
        let run = Run::default();
        assert_eq!(run.text, "");
        assert!(!run.bold);
        assert!(!run.italic);

        let p = Paragraph::default();
        assert!(p.style_id.is_none());
        assert!(p.children.is_empty());

        let doc = DocumentXml::default();
        assert_eq!(doc.width, 11906);
        assert_eq!(doc.height, 16838);
    }

    #[test]
    fn test_hyperlink() {
        let hyperlink =
            Hyperlink::new("rId1").add_run(Run::new("Click here").underline().color("0000FF"));

        assert_eq!(hyperlink.id, "rId1");
        assert_eq!(hyperlink.children.len(), 1);
        assert!(hyperlink.children[0].underline);
        assert_eq!(hyperlink.children[0].color, Some("0000FF".to_string()));
    }

    #[test]
    fn test_paragraph_with_hyperlink() {
        let p = Paragraph::new()
            .add_text("Visit ")
            .add_hyperlink(Hyperlink::new("rId1").add_run(Run::new("our website").underline()))
            .add_text(" for more info.");

        assert_eq!(p.children.len(), 3);
        match &p.children[0] {
            ParagraphChild::Run(run) => assert_eq!(run.text, "Visit "),
            _ => panic!("Expected Run child"),
        }
        match &p.children[1] {
            ParagraphChild::Hyperlink(link) => {
                assert_eq!(link.id, "rId1");
                assert_eq!(link.children.len(), 1);
                assert!(link.children[0].underline);
            }
            _ => panic!("Expected Hyperlink child"),
        }
        match &p.children[2] {
            ParagraphChild::Run(run) => assert_eq!(run.text, " for more info."),
            _ => panic!("Expected Run child"),
        }
    }

    #[test]
    fn test_paragraph_with_hyperlink_xml() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::new()
            .add_text("Click ")
            .add_hyperlink(Hyperlink::new("rId5").add_run(Run::new("here").underline()));
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:hyperlink r:id=\"rId5\">"));
        assert!(xml_str.contains("<w:u w:val=\"single\"/>"));
        assert!(xml_str.contains("here"));
        assert!(xml_str.contains("</w:hyperlink>"));
    }

    // Table tests
    #[test]
    fn test_table_basic() {
        let mut doc = DocumentXml::new();

        // Create a simple 2x2 table
        let table = Table::new()
            .with_column_widths(vec![2000, 2000])
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("A1")),
                    )
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("B1")),
                    ),
            )
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("A2")),
                    )
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("B2")),
                    ),
            );

        doc.add_table(table);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:tbl>"));
        assert!(xml_str.contains("<w:tblPr>"));
        assert!(xml_str.contains("<w:tblStyle w:val=\"TableGrid\"/>"));
        assert!(xml_str.contains("<w:tblBorders>"));
        assert!(xml_str.contains("<w:tblGrid>"));
        assert!(xml_str.contains("<w:gridCol w:w=\"2000\"/>"));
        assert!(xml_str.contains("<w:tr>"));
        assert!(xml_str.contains("<w:tc>"));
        assert!(xml_str.contains("A1"));
        assert!(xml_str.contains("B1"));
        assert!(xml_str.contains("A2"));
        assert!(xml_str.contains("B2"));
    }

    #[test]
    fn test_table_with_header() {
        let mut doc = DocumentXml::new();

        // Create a table with header row
        let table = Table::new()
            .with_column_widths(vec![2000, 2000, 2000])
            .with_header_row(true)
            .add_row(
                TableRow::new()
                    .header()
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Name"))
                            .shading("D9E2F3"),
                    )
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Age"))
                            .shading("D9E2F3"),
                    )
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("City"))
                            .shading("D9E2F3"),
                    ),
            )
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("John")),
                    )
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("30")),
                    )
                    .add_cell(
                        TableCellElement::new().add_paragraph(Paragraph::new().add_text("NYC")),
                    ),
            );

        doc.add_table(table);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:tblHeader/>"));
        assert!(xml_str.contains("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"D9E2F3\"/>"));
        assert!(xml_str.contains("Name"));
        assert!(xml_str.contains("Age"));
        assert!(xml_str.contains("City"));
        assert!(xml_str.contains("John"));
    }

    #[test]
    fn test_table_alignment() {
        let mut doc = DocumentXml::new();

        // Create a table with different alignments
        let table = Table::new()
            .with_column_widths(vec![2000, 2000, 2000])
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Left"))
                            .alignment("left"),
                    )
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Center"))
                            .alignment("center"),
                    )
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Right"))
                            .alignment("right"),
                    ),
            );

        doc.add_table(table);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:jc w:val=\"left\"/>"));
        assert!(xml_str.contains("<w:jc w:val=\"center\"/>"));
        assert!(xml_str.contains("<w:jc w:val=\"right\"/>"));
    }

    #[test]
    fn test_table_cell_width() {
        let mut doc = DocumentXml::new();

        // Create a table with custom cell widths
        let table = Table::new()
            .with_column_widths(vec![1000, 3000, 2000])
            .add_row(
                TableRow::new()
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Narrow"))
                            .width(TableWidth::Dxa(1000)),
                    )
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Wide"))
                            .width(TableWidth::Dxa(3000)),
                    )
                    .add_cell(
                        TableCellElement::new()
                            .add_paragraph(Paragraph::new().add_text("Medium"))
                            .width(TableWidth::Dxa(2000)),
                    ),
            );

        doc.add_table(table);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:gridCol w:w=\"1000\"/>"));
        assert!(xml_str.contains("<w:gridCol w:w=\"3000\"/>"));
        assert!(xml_str.contains("<w:gridCol w:w=\"2000\"/>"));
        assert!(xml_str.contains("<w:tcW w:w=\"1000\" w:type=\"dxa\"/>"));
        assert!(xml_str.contains("<w:tcW w:w=\"3000\" w:type=\"dxa\"/>"));
        assert!(xml_str.contains("<w:tcW w:w=\"2000\" w:type=\"dxa\"/>"));
    }

    #[test]
    fn test_table_borders() {
        let mut doc = DocumentXml::new();

        let table = Table::new()
            .with_column_widths(vec![2000])
            .add_row(TableRow::new().add_cell(
                TableCellElement::new().add_paragraph(Paragraph::new().add_text("Cell")),
            ));

        doc.add_table(table);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Check that all borders are present
        assert!(
            xml_str.contains("<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>")
        );
        assert!(xml_str
            .contains("<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"));
        assert!(xml_str
            .contains("<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"));
        assert!(xml_str
            .contains("<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"));
        assert!(xml_str
            .contains("<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"));
        assert!(xml_str
            .contains("<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"));
    }

    #[test]
    fn test_table_with_multiple_paragraphs_in_cell() {
        let mut doc = DocumentXml::new();

        // Create a table with multiple paragraphs in a cell
        let table = Table::new().with_column_widths(vec![2000]).add_row(
            TableRow::new().add_cell(
                TableCellElement::new()
                    .add_paragraph(Paragraph::new().add_text("First line"))
                    .add_paragraph(Paragraph::new().add_text("Second line"))
                    .add_paragraph(Paragraph::new().add_text("Third line")),
            ),
        );

        doc.add_table(table);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("First line"));
        assert!(xml_str.contains("Second line"));
        assert!(xml_str.contains("Third line"));
    }

    #[test]
    fn test_table_default_implementations() {
        let table = Table::default();
        assert!(table.rows.is_empty());
        assert!(table.column_widths.is_empty());
        assert!(!table.has_header_row);

        let row = TableRow::default();
        assert!(row.cells.is_empty());
        assert!(!row.is_header);

        let cell = TableCellElement::default();
        assert!(cell.paragraphs.is_empty());
        assert!(matches!(cell.width, TableWidth::Auto));
        assert!(cell.alignment.is_none());
        assert!(cell.shading.is_none());
    }

    #[test]
    fn test_document_with_mixed_elements() {
        let mut doc = DocumentXml::new();

        // Add a paragraph
        doc.add_paragraph(Paragraph::with_style("Heading1").add_text("Title"));

        // Add a table
        let table = Table::new().with_column_widths(vec![2000, 2000]).add_row(
            TableRow::new()
                .add_cell(TableCellElement::new().add_paragraph(Paragraph::new().add_text("A")))
                .add_cell(TableCellElement::new().add_paragraph(Paragraph::new().add_text("B"))),
        );
        doc.add_table(table);

        // Add another paragraph
        doc.add_paragraph(Paragraph::with_style("Normal").add_text("After table"));

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("Title"));
        assert!(xml_str.contains("<w:tbl>"));
        assert!(xml_str.contains("After table"));
    }

    // Image tests
    #[test]
    fn test_image_element_new() {
        let img = ImageElement::new("rId4", 914400, 457200)
            .alt_text("Test image")
            .name("test.png")
            .id(1);

        assert_eq!(img.rel_id, "rId4");
        assert_eq!(img.width_emu, 914400);
        assert_eq!(img.height_emu, 457200);
        assert_eq!(img.alt_text, "Test image");
        assert_eq!(img.name, "test.png");
    }

    #[test]
    fn test_image_element_builder() {
        let img = ImageElement::new("rId10", 2000000, 1000000)
            .alt_text("A beautiful sunset")
            .name("sunset.jpg")
            .id(5);

        assert_eq!(img.rel_id, "rId10");
        assert_eq!(img.width_emu, 2000000);
        assert_eq!(img.height_emu, 1000000);
        assert_eq!(img.alt_text, "A beautiful sunset");
        assert_eq!(img.name, "sunset.jpg");
        assert_eq!(img.id, 5);
    }

    #[test]
    fn test_image_drawing_xml_structure() {
        let mut doc = DocumentXml::new();

        let img = ImageElement::new("rId4", 914400, 457200)
            .alt_text("Test image")
            .name("test.png")
            .id(1);
        doc.add_image(img);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Verify the complete drawing structure (use partial matches to handle namespace attributes)
        assert!(xml_str.contains("<w:drawing>"));
        assert!(xml_str.contains("<wp:inline"));
        assert!(xml_str.contains("<wp:extent"));
        assert!(xml_str.contains("<wp:effectExtent"));
        assert!(xml_str.contains("<wp:docPr"));
        assert!(xml_str.contains("<wp:cNvGraphicFramePr>"));
        assert!(xml_str.contains("<a:graphic"));
        assert!(xml_str.contains("<a:graphicData"));
        assert!(xml_str.contains("<pic:pic") || xml_str.contains("pic:pic"));
        assert!(xml_str.contains("<pic:nvPicPr>"));
        assert!(xml_str.contains("<pic:blipFill>"));
        assert!(xml_str.contains("<pic:spPr>"));
        assert!(xml_str.contains("<a:xfrm>"));
        assert!(xml_str.contains("<a:prstGeom"));
    }

    #[test]
    fn test_image_wrapped_in_paragraph() {
        let mut doc = DocumentXml::new();

        let img = ImageElement::new("rId4", 914400, 457200).id(1);
        doc.add_image(img);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Images should be wrapped in paragraph and run
        assert!(xml_str.contains("<w:p>"));
        assert!(xml_str.contains("<w:r>"));
        assert!(xml_str.contains("<w:drawing>"));
        assert!(xml_str.contains("</w:drawing>"));
        assert!(xml_str.contains("</w:r>"));
        assert!(xml_str.contains("</w:p>"));
    }

    #[test]
    fn test_image_with_alt_text() {
        let mut doc = DocumentXml::new();

        let img = ImageElement::new("rId4", 914400, 457200)
            .alt_text("This is alt text for accessibility")
            .id(1);
        doc.add_image(img);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("descr=\"This is alt text for accessibility\""));
    }

    #[test]
    fn test_image_dimensions_in_xml() {
        let mut doc = DocumentXml::new();

        let img = ImageElement::new("rId4", 2000000, 1500000).id(1);
        doc.add_image(img);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Check that dimensions appear in both wp:extent and a:ext
        assert!(xml_str.contains("cx=\"2000000\""));
        assert!(xml_str.contains("cy=\"1500000\""));
    }

    #[test]
    fn test_image_aspect_ratio_lock() {
        let mut doc = DocumentXml::new();

        let img = ImageElement::new("rId4", 914400, 457200).id(1);
        doc.add_image(img);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Verify that aspect ratio is locked
        assert!(xml_str.contains("noChangeAspect=\"1\""));
    }

    #[test]
    fn test_image_relationship_id() {
        let mut doc = DocumentXml::new();

        let img = ImageElement::new("rId7", 914400, 457200).id(1);
        doc.add_image(img);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Verify the relationship ID is embedded correctly
        assert!(xml_str.contains("r:embed=\"rId7\""));
    }

    #[test]
    fn test_paragraph_shading() {
        let p = Paragraph::new().add_text("Highlighted").shading("FFFF00");
        assert_eq!(p.shading, Some("FFFF00".to_string()));

        // Verify XML generation
        let mut doc = DocumentXml::new();
        doc.add_paragraph(p);
        let xml = String::from_utf8(doc.to_xml().unwrap()).unwrap();
        assert!(xml.contains("<w:shd"));
        assert!(xml.contains("w:fill=\"FFFF00\""));
    }

    // Bookmark tests
    #[test]
    fn test_bookmark_start_struct() {
        let bookmark = BookmarkStart {
            id: 42,
            name: "_Toc_Test".to_string(),
        };
        assert_eq!(bookmark.id, 42);
        assert_eq!(bookmark.name, "_Toc_Test");
    }

    #[test]
    fn test_paragraph_with_bookmark() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::with_style("Heading1")
            .add_text("Introduction")
            .with_bookmark(1, "_Toc1_Introduction");
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains(r#"<w:bookmarkStart w:id="1" w:name="_Toc1_Introduction"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkEnd w:id="1"/>"#));
        // Bookmark should wrap the content
        assert!(xml_str.contains("Introduction"));
    }

    #[test]
    fn test_paragraph_without_bookmark() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::with_style("Normal").add_text("Regular paragraph");
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // No bookmark elements when not set
        assert!(!xml_str.contains("bookmarkStart"));
        assert!(!xml_str.contains("bookmarkEnd"));
    }

    #[test]
    fn test_bookmark_with_multiple_runs() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::new()
            .add_run(Run::new("Chapter ").bold())
            .add_run(Run::new("1").size(32))
            .with_bookmark(2, "_Toc2_Chapter1");
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains(r#"<w:bookmarkStart w:id="2" w:name="_Toc2_Chapter1"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkEnd w:id="2"/>"#));
        assert!(xml_str.contains("Chapter"));
        assert!(xml_str.contains("1"));
    }

    #[test]
    fn test_bookmark_with_style() {
        let mut doc = DocumentXml::new();
        let p = Paragraph::with_style("Heading2")
            .add_text("Section 1.1")
            .with_bookmark(3, "_Toc3_Section1_1");
        doc.add_paragraph(p);

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Verify style is present
        assert!(xml_str.contains(r#"<w:pStyle w:val="Heading2"/>"#));
        // Verify bookmark is present
        assert!(xml_str.contains(r#"<w:bookmarkStart w:id="3" w:name="_Toc3_Section1_1"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkEnd w:id="3"/>"#));
    }

    #[test]
    fn test_multiple_bookmarks_in_document() {
        let mut doc = DocumentXml::new();
        doc.add_paragraph(
            Paragraph::with_style("Heading1")
                .add_text("Chapter 1")
                .with_bookmark(1, "_Toc1_Chapter1"),
        );
        doc.add_paragraph(
            Paragraph::with_style("Heading2")
                .add_text("Section 1.1")
                .with_bookmark(2, "_Toc2_Section1_1"),
        );
        doc.add_paragraph(
            Paragraph::with_style("Heading2")
                .add_text("Section 1.2")
                .with_bookmark(3, "_Toc3_Section1_2"),
        );

        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains(r#"<w:bookmarkStart w:id="1" w:name="_Toc1_Chapter1"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkStart w:id="2" w:name="_Toc2_Section1_1"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkStart w:id="3" w:name="_Toc3_Section1_2"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkEnd w:id="1"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkEnd w:id="2"/>"#));
        assert!(xml_str.contains(r#"<w:bookmarkEnd w:id="3"/>"#));
    }

    // Header/footer tests
    #[test]
    fn test_header_footer_refs_default() {
        let refs = HeaderFooterRefs::default();
        assert!(refs.default_header_id.is_none());
        assert!(refs.default_footer_id.is_none());
        assert!(!refs.different_first_page);
    }

    #[test]
    fn test_document_with_header_footer_refs() {
        let refs = HeaderFooterRefs {
            default_header_id: Some("rId4".to_string()),
            default_footer_id: Some("rId5".to_string()),
            first_header_id: None,
            first_footer_id: None,
            different_first_page: false,
        };

        let doc = DocumentXml::new().with_header_footer(refs);
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains(r#"<w:headerReference w:type="default" r:id="rId4"/>"#));
        assert!(xml_str.contains(r#"<w:footerReference w:type="default" r:id="rId5"/>"#));
        assert!(!xml_str.contains("w:titlePg"));
    }

    #[test]
    fn test_document_with_different_first_page() {
        let refs = HeaderFooterRefs {
            default_header_id: Some("rId4".to_string()),
            first_header_id: Some("rId5".to_string()),
            default_footer_id: Some("rId6".to_string()),
            first_footer_id: Some("rId7".to_string()),
            different_first_page: true,
        };

        let doc = DocumentXml::new().with_header_footer(refs);
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains(r#"w:type="default""#));
        assert!(xml_str.contains(r#"w:type="first""#));
        assert!(xml_str.contains("<w:titlePg/>"));
    }

    #[test]
    fn test_document_with_only_default_header() {
        let refs = HeaderFooterRefs {
            default_header_id: Some("rId4".to_string()),
            first_header_id: None,
            default_footer_id: None,
            first_footer_id: None,
            different_first_page: false,
        };

        let doc = DocumentXml::new().with_header_footer(refs);
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains(r#"<w:headerReference w:type="default" r:id="rId4"/>"#));
        assert!(!xml_str.contains("footerReference"));
    }

    #[test]
    fn test_document_with_only_default_footer() {
        let refs = HeaderFooterRefs {
            default_header_id: None,
            first_header_id: None,
            default_footer_id: Some("rId6".to_string()),
            first_footer_id: None,
            different_first_page: false,
        };

        let doc = DocumentXml::new().with_header_footer(refs);
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(!xml_str.contains("headerReference"));
        assert!(xml_str.contains(r#"<w:footerReference w:type="default" r:id="rId6"/>"#));
    }

    #[test]
    fn test_header_footer_refs_order_in_sect_pr() {
        let refs = HeaderFooterRefs {
            default_header_id: Some("rId4".to_string()),
            first_header_id: Some("rId5".to_string()),
            default_footer_id: Some("rId6".to_string()),
            first_footer_id: Some("rId7".to_string()),
            different_first_page: true,
        };

        let doc = DocumentXml::new().with_header_footer(refs);
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Find the position of sectPr
        let sect_pr_start = xml_str.find("<w:sectPr>").unwrap();

        // Find positions of each element within sectPr
        let header_default_pos = xml_str[sect_pr_start..]
            .find(r#"w:type="default" r:id="rId4""#)
            .unwrap();
        let header_first_pos = xml_str[sect_pr_start..]
            .find(r#"w:type="first" r:id="rId5""#)
            .unwrap();
        let footer_default_pos = xml_str[sect_pr_start..]
            .find(r#"w:type="default" r:id="rId6""#)
            .unwrap();
        let footer_first_pos = xml_str[sect_pr_start..]
            .find(r#"w:type="first" r:id="rId7""#)
            .unwrap();
        let pg_sz_pos = xml_str[sect_pr_start..].find("<w:pgSz").unwrap();
        let title_pg_pos = xml_str[sect_pr_start..].find("<w:titlePg/>").unwrap();

        // Verify order: header refs, footer refs, pgSz, ..., titlePg
        assert!(header_default_pos < pg_sz_pos);
        assert!(header_first_pos < pg_sz_pos);
        assert!(footer_default_pos < pg_sz_pos);
        assert!(footer_first_pos < pg_sz_pos);
        assert!(title_pg_pos > pg_sz_pos);
    }

    #[test]
    fn test_suppress_header_footer_uses_empty_refs() {
        let mut doc_xml = DocumentXml::new();
        // Setup mock empty header/footer IDs
        doc_xml.empty_header_id = Some("rIdEmptyH".to_string());
        doc_xml.empty_footer_id = Some("rIdEmptyF".to_string());

        // Add a paragraph with section break and suppression enabled
        let p = Paragraph::new()
            .section_break("nextPage")
            .suppress_header_footer();
        doc_xml.add_paragraph(p);

        let xml = doc_xml.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Verify that we explicitly reference the EMPTY header/footer
        // This confirms we are preventing inheritance
        assert!(xml_str.contains("w:headerReference"));
        assert!(xml_str.contains("r:id=\"rIdEmptyH\""));
        assert!(xml_str.contains("w:footerReference"));
        assert!(xml_str.contains("r:id=\"rIdEmptyF\""));
    }
}
