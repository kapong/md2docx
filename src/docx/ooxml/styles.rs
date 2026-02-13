//! Generate word/styles.xml for DOCX

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;

/// kompleks script size in half-points (14pt = 28)
#[allow(dead_code)]
pub const DEFAULT_THAI_SIZE: u32 = 28;

/// Font configuration
#[derive(Debug, Clone, Default)]
pub struct FontConfig {
    pub default: Option<String>,
    pub code: Option<String>,
    pub normal_size: Option<u32>,
    pub normal_color: Option<String>,
    pub h1_color: Option<String>,
    pub caption_size: Option<u32>,
    pub caption_color: Option<String>,
    pub code_size: Option<u32>,
}

/// Language setting for default fonts
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Language {
    #[default]
    English,
    Thai,
}

impl Language {
    /// Get default ASCII font for this language
    pub fn default_ascii_font(&self) -> &'static str {
        match self {
            Language::English => "Calibri",
            Language::Thai => "TH Sarabun New",
        }
    }

    /// Get default complex script font for this language
    pub fn default_cs_font(&self) -> &'static str {
        match self {
            // Use TH Sarabun New for CS font even in English mode
            // This ensures mixed Thai text in English documents renders correctly
            Language::English => "TH Sarabun New",
            Language::Thai => "TH Sarabun New",
        }
    }

    /// Get default font size in half-points
    pub fn default_font_size(&self) -> u32 {
        match self {
            Language::English => 22, // 11pt
            Language::Thai => 28,    // 14pt
        }
    }

    /// Get default complex script size in half-points
    pub fn default_cs_size(&self) -> u32 {
        match self {
            Language::English => 22, // 11pt
            Language::Thai => 28,    // 14pt
        }
    }

    /// Get localized table caption prefix
    pub fn table_caption_prefix(&self) -> &'static str {
        match self {
            Language::English => "Table",
            Language::Thai => "ตารางที่",
        }
    }

    /// Get localized figure/image caption prefix
    pub fn figure_caption_prefix(&self) -> &'static str {
        match self {
            Language::English => "Figure",
            Language::Thai => "รูปที่",
        }
    }
}

/// Style type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(dead_code)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl StyleType {
    /// Convert to OOXML attribute value
    pub fn as_str(&self) -> &'static str {
        match self {
            StyleType::Paragraph => "paragraph",
            StyleType::Character => "character",
            StyleType::Table => "table",
            StyleType::Numbering => "numbering",
        }
    }
}

/// Tab stop definition
#[derive(Debug, Clone)]
pub struct TabStop {
    pub position: u32,          // Position in twips
    pub alignment: String,      // "left", "center", "right"
    pub leader: Option<String>, // "dot", "hyphen", "underscore", or None
}

impl TabStop {
    pub fn right_aligned_with_dots(position: u32) -> Self {
        Self {
            position,
            alignment: "right".to_string(),
            leader: Some("dot".to_string()),
        }
    }
}

/// Style definition
#[derive(Debug, Clone)]
#[allow(non_snake_case)]
pub struct Style {
    pub id: String,
    pub name: String,
    pub style_type: StyleType,
    pub based_on: Option<String>,
    pub next: Option<String>,
    pub ui_priority: Option<u32>, // UI priority (controls order in style gallery)
    pub font_ascii: Option<String>,
    pub font_hAnsi: Option<String>,
    pub font_cs: Option<String>, // Complex script (Thai)
    pub size: Option<u32>,       // In half-points
    pub size_cs: Option<u32>,    // Complex script size
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: Option<String>,       // Hex color without #
    pub outline_level: Option<u8>,   // For headings (0-8)
    pub spacing_before: Option<u32>, // In twips (1/20 pt)
    pub spacing_after: Option<u32>,
    pub indent_left: Option<u32>,       // In twips
    pub indent_first_line: Option<u32>, // First-line indent in twips
    pub contextual_spacing: bool, // Ignore spacing between same styles
    pub hidden: bool,
    pub semi_hidden: bool,
    pub unhide_when_used: bool,
    pub tabs: Vec<TabStop>, // Tab stops for paragraph styles
}

impl Style {
    pub fn new(id: &str, name: &str, style_type: StyleType) -> Self {
        Self {
            id: id.to_string(),
            name: name.to_string(),
            style_type,
            based_on: None,
            next: None,
            ui_priority: None,
            font_ascii: None,
            font_hAnsi: None,
            font_cs: None,
            size: None,
            size_cs: None,
            bold: false,
            italic: false,
            underline: false,
            color: None,
            outline_level: None,
            spacing_before: None,
            spacing_after: None,
            indent_left: None,
            indent_first_line: None,
            contextual_spacing: false,
            hidden: false,
            semi_hidden: false,
            unhide_when_used: false,
            tabs: Vec::new(),
        }
    }

    /// Set based-on style
    pub fn based_on(mut self, style_id: &str) -> Self {
        self.based_on = Some(style_id.to_string());
        self
    }

    /// Set next style
    pub fn next(mut self, style_id: &str) -> Self {
        self.next = Some(style_id.to_string());
        self
    }

    /// Set UI priority (controls order in style gallery)
    pub fn ui_priority(mut self, priority: u32) -> Self {
        self.ui_priority = Some(priority);
        self
    }

    /// Set font
    #[allow(non_snake_case)]
    pub fn font(mut self, ascii: &str, hAnsi: &str, cs: &str) -> Self {
        self.font_ascii = Some(ascii.to_string());
        self.font_hAnsi = Some(hAnsi.to_string());
        self.font_cs = Some(cs.to_string());
        self
    }

    /// Set size in half-points
    pub fn size(mut self, size: u32) -> Self {
        self.size = Some(size);
        self
    }

    /// Set complex script size in half-points
    pub fn size_cs(mut self, size: u32) -> Self {
        self.size_cs = Some(size);
        self
    }

    /// Set bold
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }

    /// Set italic
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }

    /// Set underline
    pub fn underline(mut self) -> Self {
        self.underline = true;
        self
    }

    /// Set color (hex without #)
    pub fn color(mut self, color: &str) -> Self {
        self.color = Some(color.to_string());
        self
    }

    /// Set outline level for headings
    pub fn outline_level(mut self, level: u8) -> Self {
        self.outline_level = Some(level);
        self
    }

    /// Set spacing before/after in twips
    pub fn spacing(mut self, before: u32, after: u32) -> Self {
        self.spacing_before = Some(before);
        self.spacing_after = Some(after);
        self
    }

    /// Set indent in twips
    pub fn indent(mut self, left: u32) -> Self {
        self.indent_left = Some(left);
        self
    }

    /// Set first-line indent in twips
    pub fn first_line_indent(mut self, first_line: u32) -> Self {
        self.indent_first_line = Some(first_line);
        self
    }

    /// Set contextual spacing
    pub fn contextual_spacing(mut self, enabled: bool) -> Self {
        self.contextual_spacing = enabled;
        self
    }

    /// Set as hidden
    #[allow(dead_code)]
    pub fn hidden(mut self) -> Self {
        self.hidden = true;
        self
    }

    /// Set as semi-hidden
    #[allow(dead_code)]
    pub fn semi_hidden(mut self) -> Self {
        self.semi_hidden = true;
        self
    }

    /// Set unhide when used
    #[allow(dead_code)]
    pub fn unhide_when_used(mut self) -> Self {
        self.unhide_when_used = true;
        self
    }

    /// Add a tab stop
    pub fn add_tab(mut self, tab: TabStop) -> Self {
        self.tabs.push(tab);
        self
    }
}

/// Styles document generator
pub(crate) struct StylesDocument {
    styles: Vec<Style>,
    lang: Language,
    font_config: Option<FontConfig>,
    /// Page width in twips (for computing header/footer tab positions)
    page_width: Option<u32>,
    /// Left margin in twips
    margin_left: Option<u32>,
    /// Right margin in twips
    margin_right: Option<u32>,
    /// Tab stops from header-footer.docx template's Header style (overrides computed)
    header_template_tabs: Option<Vec<(u32, String)>>,
    /// Tab stops from header-footer.docx template's Footer style (overrides computed)
    footer_template_tabs: Option<Vec<(u32, String)>>,
}

impl StylesDocument {
    pub fn new(lang: Language, font_config: Option<FontConfig>) -> Self {
        let mut doc = Self {
            styles: Vec::new(),
            lang,
            font_config,
            page_width: None,
            margin_left: None,
            margin_right: None,
            header_template_tabs: None,
            footer_template_tabs: None,
        };
        doc.add_default_styles();
        doc
    }

    /// Create with page dimensions so header/footer tab stops are computed correctly
    pub fn with_page_layout(
        lang: Language,
        font_config: Option<FontConfig>,
        page_width: Option<u32>,
        margin_left: Option<u32>,
        margin_right: Option<u32>,
    ) -> Self {
        let mut doc = Self {
            styles: Vec::new(),
            lang,
            font_config,
            page_width,
            margin_left,
            margin_right,
            header_template_tabs: None,
            footer_template_tabs: None,
        };
        doc.add_default_styles();
        doc
    }

    /// Set the tab stops from a header-footer.docx template's styles.
    /// When set, these override the computed tab positions for Header/Footer styles.
    pub fn set_template_tabs(
        &mut self,
        header_tabs: Vec<(u32, String)>,
        footer_tabs: Vec<(u32, String)>,
    ) {
        if !header_tabs.is_empty() {
            self.header_template_tabs = Some(header_tabs);
        }
        if !footer_tabs.is_empty() {
            self.footer_template_tabs = Some(footer_tabs);
        }
        // Rebuild styles with new tabs
        self.styles.clear();
        self.add_default_styles();
    }

    /// Compute the text area width in twips (page_width - left_margin - right_margin)
    /// Falls back to A4 defaults: 11906 - 1440 - 1440 = 9026
    fn text_area_width(&self) -> u32 {
        let width = self.page_width.unwrap_or(11906); // A4 = 210mm
        let left = self.margin_left.unwrap_or(1440); // 25.4mm ≈ 1440 twips
        let right = self.margin_right.unwrap_or(1440);
        width.saturating_sub(left).saturating_sub(right)
    }

    fn get_ascii_font(&self) -> String {
        self.font_config
            .as_ref()
            .and_then(|c| c.default.as_ref())
            .cloned()
            .unwrap_or_else(|| self.lang.default_ascii_font().to_string())
    }

    /// Get the Complex Script font — uses the configured default font so that
    /// embedded fonts are applied to Thai / CS text as well as Latin text.
    fn get_cs_font(&self) -> String {
        self.font_config
            .as_ref()
            .and_then(|c| c.default.as_ref())
            .cloned()
            .unwrap_or_else(|| self.lang.default_cs_font().to_string())
    }

    fn get_code_font(&self) -> String {
        self.font_config
            .as_ref()
            .and_then(|c| c.code.as_ref())
            .cloned()
            .unwrap_or_else(|| "Consolas".to_string())
    }

    fn get_normal_size(&self) -> u32 {
        self.font_config
            .as_ref()
            .and_then(|c| c.normal_size)
            .unwrap_or(self.lang.default_font_size())
    }

    fn get_normal_color(&self) -> String {
        self.font_config
            .as_ref()
            .and_then(|c| c.normal_color.as_ref())
            .map(|s| s.trim_start_matches('#').to_string())
            .unwrap_or_else(|| "000000".to_string())
    }

    fn get_h1_color(&self) -> String {
        self.font_config
            .as_ref()
            .and_then(|c| c.h1_color.as_ref())
            .map(|s| s.trim_start_matches('#').to_string())
            .unwrap_or_else(|| "2F5496".to_string())
    }

    fn get_caption_size(&self) -> u32 {
        self.font_config
            .as_ref()
            .and_then(|c| c.caption_size)
            .unwrap_or(match self.lang {
                Language::English => 18, // 9pt
                Language::Thai => 24,    // 12pt
            })
    }

    fn get_caption_color(&self) -> String {
        self.font_config
            .as_ref()
            .and_then(|c| c.caption_color.as_ref())
            .map(|s| s.trim_start_matches('#').to_string())
            .unwrap_or_else(|| "000000".to_string())
    }

    fn get_code_size(&self) -> u32 {
        self.font_config
            .as_ref()
            .and_then(|c| c.code_size)
            .unwrap_or(20) // 10pt
    }

    /// Add a style to the document
    pub fn add_style(&mut self, style: Style) {
        self.styles.push(style);
    }

    /// Add all required styles
    fn add_default_styles(&mut self) {
        let ascii_font = self.get_ascii_font();
        let cs_font = self.get_cs_font();
        let code_font = self.get_code_font();
        let normal_size = self.get_normal_size();
        let normal_color = self.get_normal_color();
        let h1_color = self.get_h1_color();
        let caption_size = self.get_caption_size();
        let caption_color = self.get_caption_color();
        let code_size = self.get_code_size();

        // Helper to determine Thai size
        // If we are in Thai mode, we use the normal_size (which might be larger by default or config)
        // We enforce a minimum only if using defaults, but here normal_size is already resolved.
        // So we just use normal_size.
        let normal_size_cs = normal_size;

        // Normal style (base for all paragraph styles)
        self.add_style(
            Style::new("Normal", "Normal", StyleType::Paragraph)
                .ui_priority(0)
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .color(&normal_color)
                .spacing(0, 0), // 0 before, 0pt after
        );

        // Body Text style (for regular paragraphs)
        self.add_style(
            Style::new("BodyText", "Body Text", StyleType::Paragraph)
                .ui_priority(99)
                .based_on("Normal")
                .spacing(0, 240) // 12pt after
                .first_line_indent(720), // 0.5 inch first-line indent
        );

        // Title style (cover page title)
        let (title_font, title_size, title_cs_size) = match self.lang {
            Language::English => (
                format!("{} Light", ascii_font),
                normal_size * 2 + 14,
                normal_size * 2 + 14,
            ),
            Language::Thai => (
                cs_font.clone(),
                normal_size + 44, // Relative to normal (e.g. 28 + 44 = 72)
                normal_size + 44,
            ),
        };
        self.add_style(
            Style::new("Title", "Title", StyleType::Paragraph)
                .ui_priority(10)
                .based_on("Normal")
                .font(&title_font, &title_font, &cs_font)
                .size(title_size)
                .size_cs(title_cs_size)
                .bold()
                .spacing(240, 240), // 12pt before/after
        );

        // Subtitle style (cover page subtitle)
        let subtitle_size = normal_size + 6; // 14pt if 11pt
        self.add_style(
            Style::new("Subtitle", "Subtitle", StyleType::Paragraph)
                .ui_priority(11)
                .based_on("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(subtitle_size)
                .size_cs(subtitle_size)
                .italic()
                .spacing(120, 240), // 6pt before, 12pt after
        );

        // Heading1 style
        let (h1_font, h1_size, h1_cs_size) = match self.lang {
            Language::English => (
                format!("{} Light", ascii_font),
                normal_size + 10,
                normal_size + 18,
            ),
            Language::Thai => (
                cs_font.clone(),
                normal_size + 12, // Relative: 28 + 12 = 40 (20pt)
                normal_size + 12,
            ),
        };
        self.add_style(
            Style::new("Heading1", "Heading 1", StyleType::Paragraph)
                .ui_priority(9)
                .based_on("Normal")
                .next("Normal")
                .font(&h1_font, &h1_font, &cs_font)
                .size(h1_size)
                .size_cs(h1_cs_size)
                .bold()
                .color(&h1_color) // Word blue or custom
                .outline_level(0)
                .spacing(480, 120), // 24pt before (like blank line), 6pt after
        );

        // Heading2 style
        let (h2_font, h2_size, h2_cs_size) = match self.lang {
            Language::English => (
                format!("{} Light", ascii_font),
                normal_size + 4,
                normal_size + 10,
            ),
            Language::Thai => (
                cs_font.clone(),
                normal_size + 4, // Relative: 28 + 4 = 32 (16pt)
                normal_size + 4,
            ),
        };
        self.add_style(
            Style::new("Heading2", "Heading 2", StyleType::Paragraph)
                .ui_priority(9)
                .based_on("Heading1")
                .next("Normal")
                .font(&h2_font, &h2_font, &cs_font)
                .size(h2_size)
                .size_cs(h2_cs_size)
                .bold()
                .color(&h1_color)
                .outline_level(1)
                .spacing(360, 120), // 18pt before, 6pt after
        );

        // Heading3 style
        let (h3_font, h3_size, h3_cs_size) = match self.lang {
            Language::English => (
                format!("{} Light", ascii_font),
                normal_size + 2,
                normal_size + 6,
            ),
            Language::Thai => (
                cs_font.clone(),
                normal_size, // Relative: 28 (14pt)
                normal_size,
            ),
        };
        self.add_style(
            Style::new("Heading3", "Heading 3", StyleType::Paragraph)
                .ui_priority(9)
                .based_on("Heading2")
                .next("Normal")
                .font(&h3_font, &h3_font, &cs_font)
                .size(h3_size)
                .size_cs(h3_cs_size)
                .bold()
                .color(&h1_color)
                .outline_level(2)
                .spacing(280, 80), // 14pt before, 4pt after
        );

        // Heading4 style
        let (h4_size, h4_cs_size) = match self.lang {
            Language::English => (normal_size, normal_size + 6),
            Language::Thai => (
                if normal_size > 2 {
                    normal_size - 2
                } else {
                    normal_size
                },
                if normal_size > 2 {
                    normal_size - 2
                } else {
                    normal_size
                },
            ), // 26 (13pt)
        };
        self.add_style(
            Style::new("Heading4", "Heading 4", StyleType::Paragraph)
                .ui_priority(9)
                .based_on("Heading3")
                .next("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(h4_size)
                .size_cs(h4_cs_size)
                .italic()
                .bold()
                .outline_level(3)
                .spacing(200, 80), // 10pt before, 4pt after
        );

        // Code style (code blocks)
        self.add_style(
            Style::new("Code", "Code", StyleType::Paragraph)
                .ui_priority(99)
                .based_on("Normal")
                .font(&code_font, &code_font, &code_font)
                .size(code_size)
                .size_cs(code_size)
                .spacing(120, 120) // 6pt before/after
                .contextual_spacing(true) // Merge spacing between code lines
                .indent(240), // Left indent for the block
        );

        // CodeChar style (inline code)
        self.add_style(
            Style::new("CodeChar", "Code Char", StyleType::Character)
                .ui_priority(99)
                .font(&code_font, &code_font, &code_font)
                .size(code_size)
                .size_cs(code_size)
                .color("D63384"), // Pinkish for code
        );

        // Quote style (blockquotes)
        self.add_style(
            Style::new("Quote", "Quote", StyleType::Paragraph)
                .ui_priority(29)
                .based_on("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .italic()
                .spacing(120, 120) // 6pt before/after
                .indent(720), // 0.5" indent
        );

        // Caption style (figure/table captions)
        self.add_style(
            Style::new("Caption", "Caption", StyleType::Paragraph)
                .ui_priority(35)
                .based_on("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(caption_size)
                .size_cs(caption_size)
                .color(&caption_color)
                .italic()
                .spacing(60, 240), // 3pt before, 12pt after
        );

        // TOCHeading style
        let toc_heading_size = normal_size + 6;
        self.add_style(
            Style::new("TOCHeading", "TOC Heading", StyleType::Paragraph)
                .ui_priority(39)
                .based_on("Heading1")
                .next("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(toc_heading_size)
                .size_cs(toc_heading_size)
                .bold()
                .spacing(240, 60), // 12pt before, 3pt after
        );

        // TOC styles
        // Calculate right margin position: A4 width (11906) - left margin (1440) - right margin (1440) = 9026 twips
        const TOC_TAB_POSITION: u32 = 9026;

        self.add_style(
            Style::new("TOC1", "toc 1", StyleType::Paragraph)
                .ui_priority(39)
                .based_on("Normal")
                .next("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .add_tab(TabStop::right_aligned_with_dots(TOC_TAB_POSITION))
                .spacing(0, 100), // 0 before, 5pt after
        );

        self.add_style(
            Style::new("TOC2", "toc 2", StyleType::Paragraph)
                .ui_priority(39)
                .based_on("Normal")
                .next("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .add_tab(TabStop::right_aligned_with_dots(TOC_TAB_POSITION))
                .spacing(0, 100) // 0 before, 5pt after
                .indent(440), // 0.3" indent (440 twips)
        );

        self.add_style(
            Style::new("TOC3", "toc 3", StyleType::Paragraph)
                .ui_priority(39)
                .based_on("Normal")
                .next("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .add_tab(TabStop::right_aligned_with_dots(TOC_TAB_POSITION))
                .spacing(0, 100) // 0 before, 5pt after
                .indent(880), // 0.6" indent (880 twips)
        );

        // FootnoteText style
        let footnote_size = if normal_size > 2 {
            normal_size - 2
        } else {
            normal_size
        };
        self.add_style(
            Style::new("FootnoteText", "Footnote Text", StyleType::Paragraph)
                .ui_priority(99)
                .based_on("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(footnote_size)
                .size_cs(footnote_size)
                .spacing(60, 60),
        );

        // Hyperlink style (character)
        self.add_style(
            Style::new("Hyperlink", "Hyperlink", StyleType::Character)
                .ui_priority(99)
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .color("0563C1") // Word hyperlink blue
                .underline(),
        );

        // ListParagraph style
        self.add_style(
            Style::new("ListParagraph", "List Paragraph", StyleType::Paragraph)
                .ui_priority(34)
                .based_on("Normal")
                .font(&ascii_font, &ascii_font, &cs_font)
                .size(normal_size)
                .size_cs(normal_size_cs)
                .spacing(60, 60),
        );

        // Header style (for header paragraphs)
        // Use template tabs if available, otherwise compute from page dimensions
        let text_width = self.text_area_width();
        let header_tabs: Vec<TabStop> = if let Some(ref tmpl_tabs) = self.header_template_tabs {
            tmpl_tabs.iter().map(|(pos, align)| TabStop {
                position: *pos,
                alignment: align.clone(),
                leader: None,
            }).collect()
        } else {
            vec![
                TabStop { position: text_width / 2, alignment: "center".to_string(), leader: None },
                TabStop { position: text_width, alignment: "right".to_string(), leader: None },
            ]
        };

        let mut header_style = Style::new("Header", "header", StyleType::Paragraph)
            .ui_priority(99)
            .based_on("Normal")
            .font(&ascii_font, &ascii_font, &cs_font)
            .size(normal_size)
            .size_cs(normal_size_cs);
        for tab in header_tabs {
            header_style = header_style.add_tab(tab);
        }
        self.add_style(header_style);

        // Footer style (for footer paragraphs)
        let footer_tabs: Vec<TabStop> = if let Some(ref tmpl_tabs) = self.footer_template_tabs {
            tmpl_tabs.iter().map(|(pos, align)| TabStop {
                position: *pos,
                alignment: align.clone(),
                leader: None,
            }).collect()
        } else {
            vec![
                TabStop { position: text_width / 2, alignment: "center".to_string(), leader: None },
                TabStop { position: text_width, alignment: "right".to_string(), leader: None },
            ]
        };

        let mut footer_style = Style::new("Footer", "footer", StyleType::Paragraph)
            .ui_priority(99)
            .based_on("Normal")
            .font(&ascii_font, &ascii_font, &cs_font)
            .size(normal_size)
            .size_cs(normal_size_cs);
        for tab in footer_tabs {
            footer_style = footer_style.add_tab(tab);
        }
        self.add_style(footer_style);

        // CodeFilename style (filename above code blocks)
        let code_filename_size = if normal_size > 4 { normal_size - 4 } else { 18 };
        self.add_style(
            Style::new("CodeFilename", "Code Filename", StyleType::Paragraph)
                .ui_priority(99)
                .based_on("Normal")
                .font(&code_font, &code_font, &code_font)
                .size(code_filename_size) // approx 9pt
                .size_cs(code_filename_size)
                .bold()
                .color("444444")
                .spacing(120, 0) // 6pt before
                .indent(240),
        );
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        use super::latent_styles::LatentStyles;

        let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

        // XML declaration with standalone="yes" (required by Word)
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with all required namespaces
        let mut root = BytesStart::new("w:styles");
        root.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        root.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        root.push_attribute((
            "xmlns:mc",
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ));
        root.push_attribute((
            "xmlns:w14",
            "http://schemas.microsoft.com/office/word/2010/wordml",
        ));
        root.push_attribute((
            "xmlns:w15",
            "http://schemas.microsoft.com/office/word/2012/wordml",
        ));
        root.push_attribute(("mc:Ignorable", "w14 w15"));
        writer.write_event(Event::Start(root))?;

        // Document defaults
        self.write_doc_defaults(&mut writer)?;

        // Latent styles (376 built-in Word styles catalog)
        // This must come after docDefaults and before style definitions per ECMA-376
        let latent_styles = LatentStyles::default();
        self.write_latent_styles(&mut writer, &latent_styles)?;

        // Write all styles
        for style in &self.styles {
            self.write_style(&mut writer, style)?;
        }

        // Close root
        writer.write_event(Event::End(BytesEnd::new("w:styles")))?;

        Ok(writer.into_inner().into_inner())
    }

    /// Write document defaults
    fn write_doc_defaults<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:docDefaults")))?;

        // Run properties default
        writer.write_event(Event::Start(BytesStart::new("w:rPrDefault")))?;
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        // ECMA-376 STRICT ORDERING for w:rPr in defaults:
        // 1. w:rFonts
        // 2. w:sz
        // 3. w:szCs
        // 4. w:lang
        // 5. w14:ligatures

        // 1. Default fonts
        let ascii_font = self.get_ascii_font();
        let cs_font = self.get_cs_font();
        let mut fonts = BytesStart::new("w:rFonts");
        fonts.push_attribute(("w:ascii", ascii_font.as_str()));
        fonts.push_attribute(("w:hAnsi", ascii_font.as_str()));
        fonts.push_attribute(("w:cs", cs_font.as_str()));
        writer.write_event(Event::Empty(fonts))?;

        // 2. Default size
        let normal_size = self.get_normal_size();
        let mut size = BytesStart::new("w:sz");
        size.push_attribute(("w:val", normal_size.to_string().as_str()));
        writer.write_event(Event::Empty(size))?;

        // 3. Default complex script size
        let mut size_cs = BytesStart::new("w:szCs");
        let thai_size = normal_size;
        size_cs.push_attribute(("w:val", thai_size.to_string().as_str()));
        writer.write_event(Event::Empty(size_cs))?;

        // Color (after size per ECMA-376)
        let normal_color = self.get_normal_color();
        let mut color = BytesStart::new("w:color");
        color.push_attribute(("w:val", normal_color.as_str()));
        writer.write_event(Event::Empty(color))?;

        // 4. Language setting for Thai support
        let mut lang = BytesStart::new("w:lang");
        lang.push_attribute(("w:val", "en-US"));
        lang.push_attribute(("w:eastAsia", "th-TH"));
        lang.push_attribute(("w:bidi", "th-TH"));
        writer.write_event(Event::Empty(lang))?;

        // 5. Ligatures (Thai ligature support)
        let mut ligatures = BytesStart::new("w14:ligatures");
        ligatures.push_attribute(("w14:val", "all"));
        writer.write_event(Event::Empty(ligatures))?;

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        writer.write_event(Event::End(BytesEnd::new("w:rPrDefault")))?;

        writer.write_event(Event::End(BytesEnd::new("w:docDefaults")))?;

        Ok(())
    }

    /// Write the latent styles catalog (376 built-in Word styles)
    fn write_latent_styles<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        latent: &super::latent_styles::LatentStyles,
    ) -> Result<()> {
        // Start latentStyles element with attributes
        let mut elem = BytesStart::new("w:latentStyles");
        elem.push_attribute((
            "w:defLockedState",
            if latent.def_locked_state { "1" } else { "0" },
        ));
        elem.push_attribute((
            "w:defUIPriority",
            latent.def_ui_priority.to_string().as_str(),
        ));
        elem.push_attribute((
            "w:defSemiHidden",
            if latent.def_semi_hidden { "1" } else { "0" },
        ));
        elem.push_attribute((
            "w:defUnhideWhenUsed",
            if latent.def_unhide_when_used {
                "1"
            } else {
                "0"
            },
        ));
        elem.push_attribute(("w:defQFormat", if latent.def_q_format { "1" } else { "0" }));
        elem.push_attribute(("w:count", latent.count.to_string().as_str()));
        writer.write_event(Event::Start(elem))?;

        // Write each exception
        for exc in latent.exceptions {
            let mut exc_elem = BytesStart::new("w:lsdException");
            exc_elem.push_attribute(("w:name", exc.name));

            // Only include non-default attributes
            if let Some(priority) = exc.ui_priority {
                exc_elem.push_attribute(("w:uiPriority", priority.to_string().as_str()));
            }
            if exc.semi_hidden {
                exc_elem.push_attribute(("w:semiHidden", "1"));
            }
            if exc.unhide_when_used {
                exc_elem.push_attribute(("w:unhideWhenUsed", "1"));
            }
            if exc.q_format {
                exc_elem.push_attribute(("w:qFormat", "1"));
            }

            writer.write_event(Event::Empty(exc_elem))?;
        }

        // Close latentStyles
        writer.write_event(Event::End(BytesEnd::new("w:latentStyles")))?;

        Ok(())
    }

    /// Write a single style element
    fn write_style<W: std::io::Write>(&self, writer: &mut Writer<W>, style: &Style) -> Result<()> {
        let mut style_elem = BytesStart::new("w:style");
        style_elem.push_attribute(("w:type", style.style_type.as_str()));
        style_elem.push_attribute(("w:styleId", style.id.as_str()));

        // Add w:default="1" attribute for Normal style
        if style.id == "Normal" {
            style_elem.push_attribute(("w:default", "1"));
        }

        writer.write_event(Event::Start(style_elem))?;

        // Style name
        let mut name = BytesStart::new("w:name");
        name.push_attribute(("w:val", style.name.as_str()));
        writer.write_event(Event::Empty(name))?;

        // Based on
        if let Some(ref based_on) = style.based_on {
            let mut based_on_elem = BytesStart::new("w:basedOn");
            based_on_elem.push_attribute(("w:val", based_on.as_str()));
            writer.write_event(Event::Empty(based_on_elem))?;
        }

        // Next style
        if let Some(ref next) = style.next {
            let mut next_elem = BytesStart::new("w:next");
            next_elem.push_attribute(("w:val", next.as_str()));
            writer.write_event(Event::Empty(next_elem))?;
        }

        // UI Priority (controls order in style gallery)
        if let Some(priority) = style.ui_priority {
            let mut priority_elem = BytesStart::new("w:uiPriority");
            priority_elem.push_attribute(("w:val", priority.to_string().as_str()));
            writer.write_event(Event::Empty(priority_elem))?;
        }

        // Auto-redefine (KEY: enables auto-update in Word)
        writer.write_event(Event::Empty(BytesStart::new("w:autoRedefine")))?;

        // Quick format (show in Quick Styles gallery)
        writer.write_event(Event::Empty(BytesStart::new("w:qFormat")))?;

        // Hidden flags
        if style.hidden {
            writer.write_event(Event::Empty(BytesStart::new("w:hidden")))?;
        }
        if style.semi_hidden {
            writer.write_event(Event::Empty(BytesStart::new("w:semiHidden")))?;
        }
        if style.unhide_when_used {
            writer.write_event(Event::Empty(BytesStart::new("w:unhideWhenUsed")))?;
        }

        // Paragraph properties (for paragraph styles)
        if style.style_type == StyleType::Paragraph {
            writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

            // ECMA-376 STRICT ORDERING for w:pPr:
            // 1. w:pStyle (style ID is in parent element, not here)
            // 2. w:keepNext
            // 3. w:pageBreakBefore
            // 4. w:numPr (not used in styles, only in document paragraphs)
            // 5. w:pBdr (paragraph border)
            // 6. w:shd (shading)
            // 7. w:tabs
            // 8. w:spacing
            // 9. w:ind (indentation)
            // 10. w:jc (justification - not used in styles)
            // 11. w:outlineLvl (for headings)
            // 12. w:rPr (paragraph-level run properties)
            // 13. w:sectPr (not in styles, only in document paragraphs)

            // Contextual spacing (placed before spacing per ECMA-376)
            if style.contextual_spacing {
                writer.write_event(Event::Empty(BytesStart::new("w:contextualSpacing")))?;
            }

            // 5. Paragraph border (for Code style mainly)
            if style.id == "Code" {
                writer.write_event(Event::Start(BytesStart::new("w:pBdr")))?;

                // Box border
                for side in &["w:top", "w:left", "w:bottom", "w:right"] {
                    let mut bdr = BytesStart::new(*side);
                    bdr.push_attribute(("w:val", "single"));
                    bdr.push_attribute(("w:sz", "4")); // 1/2 pt
                    bdr.push_attribute(("w:space", "4")); // 4pt padding from text
                    bdr.push_attribute(("w:color", "D0D0D0")); // Light gray border
                    writer.write_event(Event::Empty(bdr))?;
                }

                writer.write_event(Event::End(BytesEnd::new("w:pBdr")))?;
            }

            // 6. Shading (for Code style mainly - background color)
            if style.id == "Code" {
                let mut shd = BytesStart::new("w:shd");
                shd.push_attribute(("w:val", "clear"));
                shd.push_attribute(("w:color", "auto"));
                shd.push_attribute(("w:fill", "F0F0F0")); // Light gray background
                writer.write_event(Event::Empty(shd))?;
            }

            // 7. Tab stops
            if !style.tabs.is_empty() {
                writer.write_event(Event::Start(BytesStart::new("w:tabs")))?;
                for tab in &style.tabs {
                    let mut tab_elem = BytesStart::new("w:tab");
                    tab_elem.push_attribute(("w:val", tab.alignment.as_str()));
                    tab_elem.push_attribute(("w:pos", tab.position.to_string().as_str()));
                    if let Some(ref leader) = tab.leader {
                        tab_elem.push_attribute(("w:leader", leader.as_str()));
                    }
                    writer.write_event(Event::Empty(tab_elem))?;
                }
                writer.write_event(Event::End(BytesEnd::new("w:tabs")))?;
            }

            // 8. Spacing
            if style.spacing_before.is_some() || style.spacing_after.is_some() {
                let mut spacing = BytesStart::new("w:spacing");
                if let Some(before) = style.spacing_before {
                    spacing.push_attribute(("w:before", before.to_string().as_str()));
                }
                if let Some(after) = style.spacing_after {
                    spacing.push_attribute(("w:after", after.to_string().as_str()));
                }
                writer.write_event(Event::Empty(spacing))?;
            }

            // 9. Indent
            if style.indent_left.is_some() || style.indent_first_line.is_some() {
                let mut indent_elem = BytesStart::new("w:ind");
                if let Some(indent) = style.indent_left {
                    indent_elem.push_attribute(("w:left", indent.to_string().as_str()));
                }
                if let Some(first_line) = style.indent_first_line {
                    indent_elem.push_attribute(("w:firstLine", first_line.to_string().as_str()));
                }
                writer.write_event(Event::Empty(indent_elem))?;
            }

            // 11. Outline level (for headings)
            if let Some(level) = style.outline_level {
                let mut outline = BytesStart::new("w:outlineLvl");
                outline.push_attribute(("w:val", level.to_string().as_str()));
                writer.write_event(Event::Empty(outline))?;
            }

            // 12. Paragraph-level run properties with ligatures
            writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;
            let mut ligatures = BytesStart::new("w14:ligatures");
            ligatures.push_attribute(("w14:val", "all"));
            writer.write_event(Event::Empty(ligatures))?;
            writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;

            writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;
        }

        // Run properties
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

        // ECMA-376 STRICT ORDERING for w:rPr:
        // (same ordering as in document.rs)

        // 1. Fonts
        if style.font_ascii.is_some() || style.font_hAnsi.is_some() || style.font_cs.is_some() {
            let mut fonts = BytesStart::new("w:rFonts");
            if let Some(ref ascii) = style.font_ascii {
                fonts.push_attribute(("w:ascii", ascii.as_str()));
            }
            #[allow(non_snake_case)]
            if let Some(ref hAnsi) = style.font_hAnsi {
                fonts.push_attribute(("w:hAnsi", hAnsi.as_str()));
            }
            if let Some(ref cs) = style.font_cs {
                fonts.push_attribute(("w:cs", cs.as_str()));
            }
            writer.write_event(Event::Empty(fonts))?;
        }

        // 2. Bold (w:b for Western, w:bCs for Complex Script e.g. Thai)
        if style.bold {
            writer.write_event(Event::Empty(BytesStart::new("w:b")))?;
            writer.write_event(Event::Empty(BytesStart::new("w:bCs")))?;
        }

        // 3. Italic (w:i for Western, w:iCs for Complex Script e.g. Thai)
        if style.italic {
            writer.write_event(Event::Empty(BytesStart::new("w:i")))?;
            writer.write_event(Event::Empty(BytesStart::new("w:iCs")))?;
        }

        // 4. Underline
        if style.underline {
            let mut underline = BytesStart::new("w:u");
            underline.push_attribute(("w:val", "single"));
            writer.write_event(Event::Empty(underline))?;
        }

        // 5. Size
        if let Some(size) = style.size {
            let mut size_elem = BytesStart::new("w:sz");
            size_elem.push_attribute(("w:val", size.to_string().as_str()));
            writer.write_event(Event::Empty(size_elem))?;
        }

        // 6. Complex script size
        if let Some(size_cs) = style.size_cs {
            let mut size_cs_elem = BytesStart::new("w:szCs");
            size_cs_elem.push_attribute(("w:val", size_cs.to_string().as_str()));
            writer.write_event(Event::Empty(size_cs_elem))?;
        }

        // 7. Color
        if let Some(ref color) = style.color {
            let mut color_elem = BytesStart::new("w:color");
            color_elem.push_attribute(("w:val", color.as_str()));
            writer.write_event(Event::Empty(color_elem))?;
        }

        // 8. Language setting for Thai support (in all styles)
        let mut lang = BytesStart::new("w:lang");
        lang.push_attribute(("w:val", "en-US"));
        lang.push_attribute(("w:eastAsia", "th-TH"));
        lang.push_attribute(("w:bidi", "th-TH"));
        writer.write_event(Event::Empty(lang))?;

        // 9. Ligatures (Thai ligature support)
        let mut ligatures = BytesStart::new("w14:ligatures");
        ligatures.push_attribute(("w14:val", "all"));
        writer.write_event(Event::Empty(ligatures))?;

        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;

        // Close style element
        writer.write_event(Event::End(BytesEnd::new("w:style")))?;

        Ok(())
    }
}

/// Generate word/settings.xml with full Word 2013+ compatibility
pub fn generate_settings_xml() -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

    // XML declaration with standalone="yes" (required by Word)
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Root element with all Word namespaces (including 2016+ extensions)
    let mut root = BytesStart::new("w:settings");
    root.push_attribute((
        "xmlns:w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ));
    root.push_attribute((
        "xmlns:r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ));
    root.push_attribute((
        "xmlns:mc",
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
    ));
    root.push_attribute((
        "xmlns:m",
        "http://schemas.openxmlformats.org/officeDocument/2006/math",
    ));
    root.push_attribute((
        "xmlns:w14",
        "http://schemas.microsoft.com/office/word/2010/wordml",
    ));
    root.push_attribute((
        "xmlns:w15",
        "http://schemas.microsoft.com/office/word/2012/wordml",
    ));
    root.push_attribute((
        "xmlns:w16",
        "http://schemas.microsoft.com/office/word/2018/wordml",
    ));
    root.push_attribute((
        "xmlns:w16cex",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    ));
    root.push_attribute((
        "xmlns:w16cid",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    ));
    root.push_attribute((
        "xmlns:w16se",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    ));
    root.push_attribute(("xmlns:o", "urn:schemas-microsoft-com:office:office"));
    root.push_attribute(("xmlns:v", "urn:schemas-microsoft-com:vml"));
    root.push_attribute(("xmlns:w10", "urn:schemas-microsoft-com:office:word"));
    root.push_attribute((
        "xmlns:sl",
        "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    ));
    root.push_attribute(("mc:Ignorable", "w14 w15 w16se w16cid w16 w16cex"));
    writer.write_event(Event::Start(root))?;

    // Zoom (100%)
    let mut zoom = BytesStart::new("w:zoom");
    zoom.push_attribute(("w:percent", "100"));
    writer.write_event(Event::Empty(zoom))?;

    // Proof state - mark as clean to prevent spell-check popups
    let mut proof_state = BytesStart::new("w:proofState");
    proof_state.push_attribute(("w:spelling", "clean"));
    proof_state.push_attribute(("w:grammar", "clean"));
    writer.write_event(Event::Empty(proof_state))?;

    // Default tab stop (0.5")
    let mut default_tab_stop = BytesStart::new("w:defaultTabStop");
    default_tab_stop.push_attribute(("w:val", "720"));
    writer.write_event(Event::Empty(default_tab_stop))?;

    // Character spacing control (do not compress for Thai)
    let mut char_spacing = BytesStart::new("w:characterSpacingControl");
    char_spacing.push_attribute(("w:val", "doNotCompress"));
    writer.write_event(Event::Empty(char_spacing))?;

    // Footnote properties (required for proper document structure)
    writer.write_event(Event::Start(BytesStart::new("w:footnotePr")))?;
    let mut fn_sep = BytesStart::new("w:footnote");
    fn_sep.push_attribute(("w:id", "-1"));
    writer.write_event(Event::Empty(fn_sep))?;
    let mut fn_cont = BytesStart::new("w:footnote");
    fn_cont.push_attribute(("w:id", "0"));
    writer.write_event(Event::Empty(fn_cont))?;
    writer.write_event(Event::End(BytesEnd::new("w:footnotePr")))?;

    // Endnote properties
    writer.write_event(Event::Start(BytesStart::new("w:endnotePr")))?;
    let mut en_sep = BytesStart::new("w:endnote");
    en_sep.push_attribute(("w:id", "-1"));
    writer.write_event(Event::Empty(en_sep))?;
    let mut en_cont = BytesStart::new("w:endnote");
    en_cont.push_attribute(("w:id", "0"));
    writer.write_event(Event::Empty(en_cont))?;
    writer.write_event(Event::End(BytesEnd::new("w:endnotePr")))?;

    // Compatibility settings for Word 2013+ and Thai/OpenType features
    writer.write_event(Event::Start(BytesStart::new("w:compat")))?;

    // Apply breaking rules for Thai
    writer.write_event(Event::Empty(BytesStart::new("w:applyBreakingRules")))?;

    // Compatibility mode (Word 2013+ = 15)
    let mut compat_mode = BytesStart::new("w:compatSetting");
    compat_mode.push_attribute(("w:name", "compatibilityMode"));
    compat_mode.push_attribute(("w:uri", "http://schemas.microsoft.com/office/word"));
    compat_mode.push_attribute(("w:val", "15"));
    writer.write_event(Event::Empty(compat_mode))?;

    // Override table style font size and justification
    let mut override_table = BytesStart::new("w:compatSetting");
    override_table.push_attribute(("w:name", "overrideTableStyleFontSizeAndJustification"));
    override_table.push_attribute(("w:uri", "http://schemas.microsoft.com/office/word"));
    override_table.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(override_table))?;

    // Enable OpenType features (for ligatures)
    let mut opentype_features = BytesStart::new("w:compatSetting");
    opentype_features.push_attribute(("w:name", "enableOpenTypeFeatures"));
    opentype_features.push_attribute(("w:uri", "http://schemas.microsoft.com/office/word"));
    opentype_features.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(opentype_features))?;

    // Do not flip mirror indents
    let mut no_flip = BytesStart::new("w:compatSetting");
    no_flip.push_attribute(("w:name", "doNotFlipMirrorIndents"));
    no_flip.push_attribute(("w:uri", "http://schemas.microsoft.com/office/word"));
    no_flip.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(no_flip))?;

    // Differentiate multi-row table headers
    let mut diff_headers = BytesStart::new("w:compatSetting");
    diff_headers.push_attribute(("w:name", "differentiateMultirowTableHeaders"));
    diff_headers.push_attribute(("w:uri", "http://schemas.microsoft.com/office/word"));
    diff_headers.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(diff_headers))?;

    // Word 2013 track bottom hyphenation (disabled)
    let mut track_hyphen = BytesStart::new("w:compatSetting");
    track_hyphen.push_attribute(("w:name", "useWord2013TrackBottomHyphenation"));
    track_hyphen.push_attribute(("w:uri", "http://schemas.microsoft.com/office/word"));
    track_hyphen.push_attribute(("w:val", "0"));
    writer.write_event(Event::Empty(track_hyphen))?;

    writer.write_event(Event::End(BytesEnd::new("w:compat")))?;

    // RSID (Revision Session IDs) - add multiple like Word does to prevent compatibility mode
    writer.write_event(Event::Start(BytesStart::new("w:rsids")))?;
    let mut rsid_root = BytesStart::new("w:rsidRoot");
    rsid_root.push_attribute(("w:val", "00A00001"));
    writer.write_event(Event::Empty(rsid_root))?;
    // Add multiple RSIDs like Word does
    for rsid_val in ["00A00001", "004A2B7F", "006C7573", "00BF55F3", "00D4082A"] {
        let mut rsid = BytesStart::new("w:rsid");
        rsid.push_attribute(("w:val", rsid_val));
        writer.write_event(Event::Empty(rsid))?;
    }
    writer.write_event(Event::End(BytesEnd::new("w:rsids")))?;

    // Math properties (for equation support)
    writer.write_event(Event::Start(BytesStart::new("m:mathPr")))?;
    let mut math_font = BytesStart::new("m:mathFont");
    math_font.push_attribute(("m:val", "Cambria Math"));
    writer.write_event(Event::Empty(math_font))?;
    let mut brk_bin = BytesStart::new("m:brkBin");
    brk_bin.push_attribute(("m:val", "before"));
    writer.write_event(Event::Empty(brk_bin))?;
    let mut brk_bin_sub = BytesStart::new("m:brkBinSub");
    brk_bin_sub.push_attribute(("m:val", "--"));
    writer.write_event(Event::Empty(brk_bin_sub))?;
    let mut small_frac = BytesStart::new("m:smallFrac");
    small_frac.push_attribute(("m:val", "0"));
    writer.write_event(Event::Empty(small_frac))?;
    writer.write_event(Event::Empty(BytesStart::new("m:dispDef")))?;
    let mut l_margin = BytesStart::new("m:lMargin");
    l_margin.push_attribute(("m:val", "0"));
    writer.write_event(Event::Empty(l_margin))?;
    let mut r_margin = BytesStart::new("m:rMargin");
    r_margin.push_attribute(("m:val", "0"));
    writer.write_event(Event::Empty(r_margin))?;
    let mut def_jc = BytesStart::new("m:defJc");
    def_jc.push_attribute(("m:val", "centerGroup"));
    writer.write_event(Event::Empty(def_jc))?;
    let mut wrap_indent = BytesStart::new("m:wrapIndent");
    wrap_indent.push_attribute(("m:val", "1440"));
    writer.write_event(Event::Empty(wrap_indent))?;
    let mut int_lim = BytesStart::new("m:intLim");
    int_lim.push_attribute(("m:val", "subSup"));
    writer.write_event(Event::Empty(int_lim))?;
    let mut nary_lim = BytesStart::new("m:naryLim");
    nary_lim.push_attribute(("m:val", "undOvr"));
    writer.write_event(Event::Empty(nary_lim))?;
    writer.write_event(Event::End(BytesEnd::new("m:mathPr")))?;

    // Theme font languages
    let mut theme_font_lang = BytesStart::new("w:themeFontLang");
    theme_font_lang.push_attribute(("w:val", "en-US"));
    theme_font_lang.push_attribute(("w:eastAsia", "th-TH"));
    theme_font_lang.push_attribute(("w:bidi", "th-TH"));
    writer.write_event(Event::Empty(theme_font_lang))?;

    // Color scheme mapping (theme colors)
    let mut clr_scheme = BytesStart::new("w:clrSchemeMapping");
    clr_scheme.push_attribute(("w:bg1", "light1"));
    clr_scheme.push_attribute(("w:t1", "dark1"));
    clr_scheme.push_attribute(("w:bg2", "light2"));
    clr_scheme.push_attribute(("w:t2", "dark2"));
    clr_scheme.push_attribute(("w:accent1", "accent1"));
    clr_scheme.push_attribute(("w:accent2", "accent2"));
    clr_scheme.push_attribute(("w:accent3", "accent3"));
    clr_scheme.push_attribute(("w:accent4", "accent4"));
    clr_scheme.push_attribute(("w:accent5", "accent5"));
    clr_scheme.push_attribute(("w:accent6", "accent6"));
    clr_scheme.push_attribute(("w:hyperlink", "hyperlink"));
    clr_scheme.push_attribute(("w:followedHyperlink", "followedHyperlink"));
    writer.write_event(Event::Empty(clr_scheme))?;

    // Update fields on open (for TOC)
    let mut update_fields = BytesStart::new("w:updateFields");
    update_fields.push_attribute(("w:val", "true"));
    writer.write_event(Event::Empty(update_fields))?;

    // Decimal symbol and list separator (locale)
    let mut decimal = BytesStart::new("w:decimalSymbol");
    decimal.push_attribute(("w:val", "."));
    writer.write_event(Event::Empty(decimal))?;
    let mut list_sep = BytesStart::new("w:listSeparator");
    list_sep.push_attribute(("w:val", ","));
    writer.write_event(Event::Empty(list_sep))?;

    // Document ID (Word 2010+)
    let mut doc_id_14 = BytesStart::new("w14:docId");
    doc_id_14.push_attribute(("w14:val", "00A00001"));
    writer.write_event(Event::Empty(doc_id_14))?;

    // Document ID (Word 2012+) - GUID format
    let mut doc_id_15 = BytesStart::new("w15:docId");
    doc_id_15.push_attribute(("w15:val", "{00A00001-0000-0000-0000-000000000001}"));
    writer.write_event(Event::Empty(doc_id_15))?;

    // Close root
    writer.write_event(Event::End(BytesEnd::new("w:settings")))?;

    Ok(writer.into_inner().into_inner())
}

/// Generate word/fontTable.xml
///
/// If `embedded_fonts` is provided, adds `w:embedRegular`/`w:embedBold`/etc. references
/// to the font entries for fonts that have embedded files.
pub fn generate_font_table_xml(
    _lang: Language,
    embedded_fonts: Option<&[crate::docx::font_embed::EmbeddedFont]>,
) -> Result<Vec<u8>> {
    use crate::docx::font_embed::{group_by_font_name, EmbeddedFont};

    let embedded_map = embedded_fonts
        .map(|fonts| group_by_font_name(fonts))
        .unwrap_or_default();

    let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

    // XML declaration with standalone="yes" (required by Word)
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Root element
    let mut root = BytesStart::new("w:fonts");
    root.push_attribute((
        "xmlns:w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ));
    root.push_attribute((
        "xmlns:r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ));
    writer.write_event(Event::Start(root))?;

    /// Helper: write a single font entry with optional embedding
    fn write_font_entry<W: std::io::Write>(
        writer: &mut Writer<W>,
        name: &str,
        panose: Option<&str>,
        charset: &str,
        family: &str,
        pitch: &str,
        embeds: Option<&Vec<&EmbeddedFont>>,
        sig: Option<(u32, u32, u32, u32, u32, u32)>,
    ) -> Result<()> {
        let mut font_elem = BytesStart::new("w:font");
        font_elem.push_attribute(("w:name", name));
        writer.write_event(Event::Start(font_elem))?;

        if let Some(panose_val) = panose {
            let mut p = BytesStart::new("w:panose1");
            p.push_attribute(("w:val", panose_val));
            writer.write_event(Event::Empty(p))?;
        }

        let mut cs = BytesStart::new("w:charset");
        cs.push_attribute(("w:val", charset));
        writer.write_event(Event::Empty(cs))?;

        let mut fam = BytesStart::new("w:family");
        fam.push_attribute(("w:val", family));
        writer.write_event(Event::Empty(fam))?;

        let mut p = BytesStart::new("w:pitch");
        p.push_attribute(("w:val", pitch));
        writer.write_event(Event::Empty(p))?;

        // Write Unicode/CodePage signature if available
        if let Some((usb0, usb1, usb2, usb3, csb0, csb1)) = sig {
            let mut sig_elem = BytesStart::new("w:sig");
            sig_elem.push_attribute(("w:usb0", format!("{:08X}", usb0).as_str()));
            sig_elem.push_attribute(("w:usb1", format!("{:08X}", usb1).as_str()));
            sig_elem.push_attribute(("w:usb2", format!("{:08X}", usb2).as_str()));
            sig_elem.push_attribute(("w:usb3", format!("{:08X}", usb3).as_str()));
            sig_elem.push_attribute(("w:csb0", format!("{:08X}", csb0).as_str()));
            sig_elem.push_attribute(("w:csb1", format!("{:08X}", csb1).as_str()));
            writer.write_event(Event::Empty(sig_elem))?;
        }

        // Write embed references if this font has embedded files
        if let Some(font_embeds) = embeds {
            for embed in font_embeds {
                let mut embed_elem = BytesStart::new(embed.variant.xml_element());
                embed_elem.push_attribute(("r:id", embed.rel_id.as_str()));
                embed_elem.push_attribute(("w:fontKey", format!("{{{}}}", embed.guid).as_str()));
                writer.write_event(Event::Empty(embed_elem))?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:font")))?;
        Ok(())
    }

    write_font_entry(
        &mut writer,
        "Calibri",
        Some("020F0502020204030204"),
        "00",
        "swiss",
        "variable",
        embedded_map.get("Calibri"),
        Some((0xE4002EFF, 0xC000247B, 0x00000009, 0x00000000, 0x000001FF, 0x00000000)),
    )?;

    write_font_entry(
        &mut writer,
        "Calibri Light",
        Some("020F0302020204030204"),
        "00",
        "swiss",
        "variable",
        embedded_map.get("Calibri Light"),
        Some((0xE4002EFF, 0xC000247B, 0x00000009, 0x00000000, 0x000001FF, 0x00000000)),
    )?;

    write_font_entry(
        &mut writer,
        "Consolas",
        Some("020B0609020204030204"),
        "00",
        "modern",
        "fixed",
        embedded_map.get("Consolas"),
        Some((0xE00006FF, 0x0000FCFF, 0x00000001, 0x00000000, 0x0000019F, 0x00000000)),
    )?;

    write_font_entry(
        &mut writer,
        "TH Sarabun New",
        Some("020B0500040200020003"),
        "DE",
        "swiss",
        "variable",
        embedded_map.get("TH Sarabun New"),
        Some((0xA100006F, 0x5000205A, 0x00000000, 0x00000000, 0x60010183, 0x80000000)),
    )?;

    // Write any additionally embedded fonts not already in the standard list
    let standard_fonts = ["Calibri", "Calibri Light", "Consolas", "TH Sarabun New"];
    for (font_name, embeds) in &embedded_map {
        if !standard_fonts.contains(&font_name.as_str()) {
            // Extract metrics from the first font variant in the group
            let metrics = embeds.iter().find_map(|e| e.metrics.as_ref());
            let (panose, charset, family, pitch, sig) = if let Some(m) = metrics {
                (
                    Some(m.panose.as_str()),
                    m.charset.as_str(),
                    m.family.as_str(),
                    m.pitch.as_str(),
                    Some((m.usb0, m.usb1, m.usb2, m.usb3, m.csb0, m.csb1)),
                )
            } else {
                (None, "00", "auto", "variable", None)
            };
            write_font_entry(
                &mut writer,
                font_name,
                panose,
                charset,
                family,
                pitch,
                Some(embeds),
                sig,
            )?;
        }
    }

    // Close root
    writer.write_event(Event::End(BytesEnd::new("w:fonts")))?;

    Ok(writer.into_inner().into_inner())
}

/// Generate word/webSettings.xml
pub fn generate_web_settings_xml() -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

    // XML declaration
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Root element with namespaces (minimal web settings)
    let mut root = BytesStart::new("w:webSettings");
    root.push_attribute((
        "xmlns:w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ));
    root.push_attribute((
        "xmlns:r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ));
    root.push_attribute((
        "xmlns:mc",
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
    ));
    root.push_attribute((
        "xmlns:w14",
        "http://schemas.microsoft.com/office/word/2010/wordml",
    ));
    root.push_attribute((
        "xmlns:w15",
        "http://schemas.microsoft.com/office/word/2012/wordml",
    ));
    root.push_attribute(("mc:Ignorable", "w14 w15"));
    writer.write_event(Event::Empty(root))?;

    Ok(writer.into_inner().into_inner())
}

/// Generate word/theme/theme1.xml
pub fn generate_theme_xml() -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

    // XML declaration
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Root element
    let mut root = BytesStart::new("a:theme");
    root.push_attribute((
        "xmlns:a",
        "http://schemas.openxmlformats.org/drawingml/2006/main",
    ));
    root.push_attribute(("name", "Office Theme"));
    writer.write_event(Event::Start(root))?;

    // Theme elements
    writer.write_event(Event::Start(BytesStart::new("a:themeElements")))?;

    // Color scheme
    let mut clr_scheme = BytesStart::new("a:clrScheme");
    clr_scheme.push_attribute(("name", "Office"));
    writer.write_event(Event::Start(clr_scheme))?;

    // Dark 1 (window text)
    writer.write_event(Event::Start(BytesStart::new("a:dk1")))?;
    let mut sys_clr = BytesStart::new("a:sysClr");
    sys_clr.push_attribute(("val", "windowText"));
    sys_clr.push_attribute(("lastClr", "000000"));
    writer.write_event(Event::Empty(sys_clr))?;
    writer.write_event(Event::End(BytesEnd::new("a:dk1")))?;

    // Light 1 (window background)
    writer.write_event(Event::Start(BytesStart::new("a:lt1")))?;
    let mut sys_clr = BytesStart::new("a:sysClr");
    sys_clr.push_attribute(("val", "window"));
    sys_clr.push_attribute(("lastClr", "FFFFFF"));
    writer.write_event(Event::Empty(sys_clr))?;
    writer.write_event(Event::End(BytesEnd::new("a:lt1")))?;

    // Dark 2
    writer.write_event(Event::Start(BytesStart::new("a:dk2")))?;
    let mut srgb = BytesStart::new("a:srgbClr");
    srgb.push_attribute(("val", "0E2841"));
    writer.write_event(Event::Empty(srgb))?;
    writer.write_event(Event::End(BytesEnd::new("a:dk2")))?;

    // Light 2
    writer.write_event(Event::Start(BytesStart::new("a:lt2")))?;
    let mut srgb = BytesStart::new("a:srgbClr");
    srgb.push_attribute(("val", "E8E8E8"));
    writer.write_event(Event::Empty(srgb))?;
    writer.write_event(Event::End(BytesEnd::new("a:lt2")))?;

    // Accent 1-6
    let accent_colors = ["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E"];
    for (i, color) in accent_colors.iter().enumerate() {
        writer.write_event(Event::Start(BytesStart::new(format!("a:accent{}", i + 1))))?;
        let mut srgb = BytesStart::new("a:srgbClr");
        srgb.push_attribute(("val", *color));
        writer.write_event(Event::Empty(srgb))?;
        writer.write_event(Event::End(BytesEnd::new(format!("a:accent{}", i + 1))))?;
    }

    // Hyperlink
    writer.write_event(Event::Start(BytesStart::new("a:hlink")))?;
    let mut srgb = BytesStart::new("a:srgbClr");
    srgb.push_attribute(("val", "467886"));
    writer.write_event(Event::Empty(srgb))?;
    writer.write_event(Event::End(BytesEnd::new("a:hlink")))?;

    // Followed hyperlink
    writer.write_event(Event::Start(BytesStart::new("a:folHlink")))?;
    let mut srgb = BytesStart::new("a:srgbClr");
    srgb.push_attribute(("val", "96607D"));
    writer.write_event(Event::Empty(srgb))?;
    writer.write_event(Event::End(BytesEnd::new("a:folHlink")))?;

    writer.write_event(Event::End(BytesEnd::new("a:clrScheme")))?;

    // Font scheme (minimal)
    let mut font_scheme = BytesStart::new("a:fontScheme");
    font_scheme.push_attribute(("name", "Office"));
    writer.write_event(Event::Start(font_scheme))?;

    // Major font (headings)
    writer.write_event(Event::Start(BytesStart::new("a:majorFont")))?;
    let mut latin = BytesStart::new("a:latin");
    latin.push_attribute(("typeface", "Calibri Light"));
    writer.write_event(Event::Empty(latin))?;
    let mut ea = BytesStart::new("a:ea");
    ea.push_attribute(("typeface", ""));
    writer.write_event(Event::Empty(ea))?;
    let mut cs = BytesStart::new("a:cs");
    cs.push_attribute(("typeface", ""));
    writer.write_event(Event::Empty(cs))?;
    writer.write_event(Event::End(BytesEnd::new("a:majorFont")))?;

    // Minor font (body)
    writer.write_event(Event::Start(BytesStart::new("a:minorFont")))?;
    let mut latin = BytesStart::new("a:latin");
    latin.push_attribute(("typeface", "Calibri"));
    writer.write_event(Event::Empty(latin))?;
    let mut ea = BytesStart::new("a:ea");
    ea.push_attribute(("typeface", ""));
    writer.write_event(Event::Empty(ea))?;
    let mut cs = BytesStart::new("a:cs");
    cs.push_attribute(("typeface", ""));
    writer.write_event(Event::Empty(cs))?;
    writer.write_event(Event::End(BytesEnd::new("a:minorFont")))?;

    writer.write_event(Event::End(BytesEnd::new("a:fontScheme")))?;

    // Format scheme (minimal)
    let mut fmt_scheme = BytesStart::new("a:fmtScheme");
    fmt_scheme.push_attribute(("name", "Office"));
    writer.write_event(Event::Start(fmt_scheme))?;

    // Fill style list (empty)
    writer.write_event(Event::Start(BytesStart::new("a:fillStyleLst")))?;
    let solid_fill = BytesStart::new("a:solidFill");
    writer.write_event(Event::Start(solid_fill))?;
    let mut scheme_clr = BytesStart::new("a:schemeClr");
    scheme_clr.push_attribute(("val", "phClr"));
    writer.write_event(Event::Empty(scheme_clr))?;
    writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    writer.write_event(Event::End(BytesEnd::new("a:fillStyleLst")))?;

    // Line style list (empty)
    writer.write_event(Event::Start(BytesStart::new("a:lnStyleLst")))?;
    let mut ln = BytesStart::new("a:ln");
    ln.push_attribute(("w", "6350"));
    ln.push_attribute(("cap", "flat"));
    ln.push_attribute(("cmpd", "sng"));
    ln.push_attribute(("algn", "ctr"));
    writer.write_event(Event::Start(ln))?;
    let solid_fill = BytesStart::new("a:solidFill");
    writer.write_event(Event::Start(solid_fill))?;
    let mut scheme_clr = BytesStart::new("a:schemeClr");
    scheme_clr.push_attribute(("val", "phClr"));
    writer.write_event(Event::Empty(scheme_clr))?;
    writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    writer.write_event(Event::End(BytesEnd::new("a:ln")))?;
    writer.write_event(Event::End(BytesEnd::new("a:lnStyleLst")))?;

    // Effect style list (empty)
    writer.write_event(Event::Start(BytesStart::new("a:effectStyleLst")))?;
    writer.write_event(Event::Start(BytesStart::new("a:effectStyle")))?;
    writer.write_event(Event::Start(BytesStart::new("a:effectLst")))?;
    writer.write_event(Event::End(BytesEnd::new("a:effectLst")))?;
    writer.write_event(Event::End(BytesEnd::new("a:effectStyle")))?;
    writer.write_event(Event::End(BytesEnd::new("a:effectStyleLst")))?;

    // Background fill style list (empty)
    writer.write_event(Event::Start(BytesStart::new("a:bgFillStyleLst")))?;
    let solid_fill = BytesStart::new("a:solidFill");
    writer.write_event(Event::Start(solid_fill))?;
    let mut scheme_clr = BytesStart::new("a:schemeClr");
    scheme_clr.push_attribute(("val", "phClr"));
    writer.write_event(Event::Empty(scheme_clr))?;
    writer.write_event(Event::End(BytesEnd::new("a:solidFill")))?;
    writer.write_event(Event::End(BytesEnd::new("a:bgFillStyleLst")))?;

    writer.write_event(Event::End(BytesEnd::new("a:fmtScheme")))?;

    writer.write_event(Event::End(BytesEnd::new("a:themeElements")))?;

    // Object defaults (optional)
    writer.write_event(Event::Empty(BytesStart::new("a:objectDefaults")))?;

    // Extra color scheme list (optional)
    writer.write_event(Event::Empty(BytesStart::new("a:extraClrSchemeLst")))?;

    writer.write_event(Event::End(BytesEnd::new("a:theme")))?;

    Ok(writer.into_inner().into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_styles_document_english() {
        let doc = StylesDocument::new(Language::English, None);
        assert_eq!(doc.styles.len(), 22); // All required styles (including TOCHeading, BodyText, CodeFilename, Header, and Footer)

        // Check Normal style
        let normal = doc.styles.iter().find(|s| s.id == "Normal").unwrap();
        assert_eq!(normal.style_type, StyleType::Paragraph);
        assert_eq!(normal.font_ascii, Some("Calibri".to_string()));
        assert_eq!(normal.size, Some(22)); // 11pt

        // Check BodyText style
        let body_text = doc.styles.iter().find(|s| s.id == "BodyText").unwrap();
        assert_eq!(body_text.based_on, Some("Normal".to_string()));
        assert_eq!(body_text.spacing_after, Some(240));

        // Check Heading1
        let h1 = doc.styles.iter().find(|s| s.id == "Heading1").unwrap();
        assert_eq!(h1.bold, true);
        assert_eq!(h1.color, Some("2F5496".to_string()));
        assert_eq!(h1.outline_level, Some(0));

        // Check TOCHeading
        let toc_heading = doc.styles.iter().find(|s| s.id == "TOCHeading").unwrap();
        assert_eq!(toc_heading.based_on, Some("Heading1".to_string()));
        assert_eq!(toc_heading.size, Some(28)); // 14pt
        assert_eq!(toc_heading.bold, true);
        assert_eq!(toc_heading.spacing_before, Some(240));
        assert_eq!(toc_heading.spacing_after, Some(60));

        // Check TOC1 has tabs
        let toc1 = doc.styles.iter().find(|s| s.id == "TOC1").unwrap();
        assert_eq!(toc1.tabs.len(), 1);
        assert_eq!(toc1.tabs[0].alignment, "right");
        assert_eq!(toc1.tabs[0].leader, Some("dot".to_string()));
        assert_eq!(toc1.tabs[0].position, 9026);

        // Check Header style has correct tab stops
        let header = doc.styles.iter().find(|s| s.id == "Header").unwrap();
        assert_eq!(header.style_type, StyleType::Paragraph);
        assert_eq!(header.based_on, Some("Normal".to_string()));
        assert_eq!(header.tabs.len(), 2);
        assert_eq!(header.tabs[0].position, 4513);
        assert_eq!(header.tabs[0].alignment, "center");
        assert_eq!(header.tabs[0].leader, None);
        assert_eq!(header.tabs[1].position, 9026);
        assert_eq!(header.tabs[1].alignment, "right");
        assert_eq!(header.tabs[1].leader, None);

        // Check Footer style has correct tab stops
        let footer = doc.styles.iter().find(|s| s.id == "Footer").unwrap();
        assert_eq!(footer.style_type, StyleType::Paragraph);
        assert_eq!(footer.based_on, Some("Normal".to_string()));
        assert_eq!(footer.tabs.len(), 2);
        assert_eq!(footer.tabs[0].position, 4513);
        assert_eq!(footer.tabs[0].alignment, "center");
        assert_eq!(footer.tabs[0].leader, None);
        assert_eq!(footer.tabs[1].position, 9026);
        assert_eq!(footer.tabs[1].alignment, "right");
        assert_eq!(footer.tabs[1].leader, None);
    }

    #[test]
    fn test_styles_document_thai() {
        let doc = StylesDocument::new(Language::Thai, None);
        assert_eq!(doc.styles.len(), 22); // All required styles (including TOCHeading, BodyText, CodeFilename, Header, and Footer)

        // Check Normal style uses Thai font
        let normal = doc.styles.iter().find(|s| s.id == "Normal").unwrap();
        assert_eq!(normal.font_ascii, Some("TH Sarabun New".to_string()));
        assert_eq!(normal.font_cs, Some("TH Sarabun New".to_string()));
        assert_eq!(normal.size, Some(28)); // 14pt for Thai

        // Check Heading1
        let h1 = doc.styles.iter().find(|s| s.id == "Heading1").unwrap();
        assert_eq!(h1.size_cs, Some(40)); // 20pt for Thai

        // Check TOCHeading
        let toc_heading = doc.styles.iter().find(|s| s.id == "TOCHeading").unwrap();
        assert_eq!(toc_heading.based_on, Some("Heading1".to_string()));
        assert_eq!(toc_heading.size, Some(34)); // 17pt
        assert_eq!(toc_heading.bold, true);
    }

    #[test]
    fn test_style_builder() {
        let style = Style::new("Test", "Test Style", StyleType::Paragraph)
            .based_on("Normal")
            .next("Normal")
            .font("Arial", "Arial", "Arial")
            .size(24)
            .size_cs(28)
            .bold()
            .italic()
            .color("FF0000")
            .outline_level(1)
            .spacing(100, 200)
            .indent(50);

        assert_eq!(style.id, "Test");
        assert_eq!(style.based_on, Some("Normal".to_string()));
        assert_eq!(style.bold, true);
        assert_eq!(style.italic, true);
        assert_eq!(style.size, Some(24));
        assert_eq!(style.outline_level, Some(1));
    }

    #[test]
    fn test_xml_structure() {
        let doc = StylesDocument::new(Language::English, None);
        let xml = doc.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Verify XML structure
        assert!(xml_str.contains("<?xml version"));
        assert!(xml_str
            .contains("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""));
        assert!(xml_str
            .contains("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""));
        assert!(
            xml_str.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"")
        );
        assert!(
            xml_str.contains("xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"")
        );
        assert!(xml_str.contains("mc:Ignorable=\"w14 w15\""));
        assert!(xml_str.contains("<w:docDefaults>"));
        assert!(xml_str.contains("<w:rPrDefault>"));
        assert!(xml_str.contains("<w:rFonts"));
        assert!(xml_str.contains("<w:sz w:val=\"22\"")); // 11pt
        assert!(xml_str.contains("<w:szCs w:val=\"22\""));

        // Verify language settings for Thai support
        assert!(xml_str.contains("<w:lang"));
        assert!(xml_str.contains("w:val=\"en-US\""));
        assert!(xml_str.contains("w:eastAsia=\"th-TH\""));
        assert!(xml_str.contains("w:bidi=\"th-TH\""));

        // Verify all required styles are present
        let required_styles = [
            "Normal",
            "BodyText",
            "Title",
            "Subtitle",
            "Heading1",
            "Heading2",
            "Heading3",
            "Heading4",
            "Code",
            "CodeChar",
            "Quote",
            "Caption",
            "TOCHeading",
            "TOC1",
            "TOC2",
            "TOC3",
            "FootnoteText",
            "Hyperlink",
            "ListParagraph",
        ];

        for style_id in &required_styles {
            assert!(
                xml_str.contains(&format!("w:styleId=\"{}\"", style_id)),
                "Style {} not found in XML",
                style_id
            );
        }

        // Verify auto-update and quick format are present for all styles
        let auto_redefine_count = xml_str.matches("<w:autoRedefine/>").count();
        assert!(
            auto_redefine_count >= 19,
            "Expected at least 19 autoRedefine elements, found {}",
            auto_redefine_count
        );

        let qformat_count = xml_str.matches("<w:qFormat/>").count();
        assert!(
            qformat_count >= 19,
            "Expected at least 19 qFormat elements, found {}",
            qformat_count
        );
    }

    #[test]
    fn test_generate_settings_xml() {
        let xml = generate_settings_xml().unwrap();
        assert!(!xml.is_empty());

        let xml_str = String::from_utf8(xml).unwrap();
        assert!(xml_str.contains("<w:settings"));
        assert!(xml_str.contains("xmlns:mc"));
        assert!(xml_str.contains("xmlns:w14"));
        assert!(xml_str.contains("xmlns:w15"));
        assert!(xml_str.contains("<w:compat>"));
        assert!(xml_str.contains("<w:applyBreakingRules/>"));
        assert!(xml_str.contains("<w:characterSpacingControl w:val=\"doNotCompress\"/>"));
        assert!(xml_str.contains("<w:themeFontLang"));
        assert!(xml_str.contains("th-TH"));
        assert!(xml_str.contains("<w:updateFields w:val=\"true\"/>"));
    }

    #[test]
    fn test_generate_font_table_xml() {
        let xml = generate_font_table_xml(Language::Thai, None).unwrap();
        assert!(!xml.is_empty());

        let xml_str = String::from_utf8(xml).unwrap();
        assert!(xml_str.contains("<w:fonts"));
        assert!(xml_str.contains("Calibri"));
        assert!(xml_str.contains("Consolas"));
        assert!(xml_str.contains("TH Sarabun New"));
    }

    #[test]
    fn test_language_defaults() {
        assert_eq!(Language::English.default_ascii_font(), "Calibri");
        assert_eq!(Language::Thai.default_ascii_font(), "TH Sarabun New");
        assert_eq!(Language::English.default_font_size(), 22); // 11pt
        assert_eq!(Language::Thai.default_font_size(), 28); // 14pt
    }

    #[test]
    fn test_style_type_as_str() {
        assert_eq!(StyleType::Paragraph.as_str(), "paragraph");
        assert_eq!(StyleType::Character.as_str(), "character");
        assert_eq!(StyleType::Table.as_str(), "table");
        assert_eq!(StyleType::Numbering.as_str(), "numbering");
    }

    #[test]
    fn test_language_table_caption_prefix() {
        assert_eq!(Language::English.table_caption_prefix(), "Table");
        assert_eq!(Language::Thai.table_caption_prefix(), "ตารางที่");
    }

    #[test]
    fn test_language_figure_caption_prefix() {
        assert_eq!(Language::English.figure_caption_prefix(), "Figure");
        assert_eq!(Language::Thai.figure_caption_prefix(), "รูปที่");
    }
}
