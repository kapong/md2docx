//! Table of Contents generation for DOCX documents

use crate::docx::ooxml::{DocElement, Paragraph, Run};

/// TOC configuration
#[derive(Debug, Clone)]
pub struct TocConfig {
    pub enabled: bool,
    pub depth: u8,         // 1-6, how many heading levels to include (default 2)
    pub title: String,     // "Table of Contents" or localized
    pub after_cover: bool, // If true, TOC comes after cover content
}

impl Default for TocConfig {
    fn default() -> Self {
        Self {
            enabled: true,
            depth: 2,
            title: "Table of Contents".to_string(),
            after_cover: true,
        }
    }
}

/// A collected heading for TOC
#[derive(Debug, Clone)]
pub struct TocEntry {
    pub text: String,        // Heading text
    pub level: u8,           // 1-6
    pub bookmark_id: String, // Bookmark name for linking (e.g., "_Toc1_Introduction")
}

/// Collects headings during document build and generates TOC
#[derive(Debug, Default)]
pub struct TocBuilder {
    entries: Vec<TocEntry>,
    next_id: u32,
}

impl TocBuilder {
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
            next_id: 0,
        }
    }

    /// Add a heading and return the bookmark ID to use
    /// If explicit_id is provided (from {#id} syntax), use it; otherwise generate one
    pub fn add_heading(&mut self, level: u8, text: &str, explicit_id: Option<&str>) -> String {
        let bookmark_id = if let Some(id) = explicit_id {
            id.to_string()
        } else {
            self.generate_bookmark_id(text)
        };

        self.entries.push(TocEntry {
            text: text.to_string(),
            level,
            bookmark_id: bookmark_id.clone(),
        });

        bookmark_id
    }

    /// Get all collected entries
    pub fn entries(&self) -> &[TocEntry] {
        &self.entries
    }

    /// Check if there are any entries
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Generate a sanitized bookmark ID from text
    fn generate_bookmark_id(&mut self, text: &str) -> String {
        self.next_id += 1;
        // Sanitize: keep only alphanumeric and underscore, max 40 chars
        let sanitized: String = text
            .chars()
            .filter(|c| c.is_alphanumeric() || *c == '_' || *c == ' ')
            .map(|c| if c == ' ' { '_' } else { c })
            .take(40)
            .collect();
        format!("_Toc{}_{}", self.next_id, sanitized)
    }

    /// Generate TOC as document elements
    /// Returns paragraphs for: TOC title + TOC field with page numbers + section break
    pub fn generate_toc(&self, config: &TocConfig) -> Vec<DocElement> {
        if !config.enabled || self.entries.is_empty() {
            return vec![];
        }

        let mut elements = Vec::new();

        // 1. TOC Title paragraph (style: TOCHeading)
        let title_para = Paragraph::with_style("TOCHeading")
            .add_text(&config.title)
            .spacing(0, 0)
            .line_spacing(240, "auto");
        elements.push(DocElement::Paragraph(Box::new(title_para)));

        // 2. TOC Field begin - Word will auto-generate entries with page numbers
        // The field code: TOC \o "1-2" \h \z \u
        // \o "1-2" = outline levels 1-2
        // \h = hyperlink entries
        // \z = preserve tab leader
        // \u = use paragraph styles
        let toc_field_begin = Paragraph::new()
            .spacing(0, 0)
            .line_spacing(240, "auto")
            .add_run(Run::new("").with_field_char("begin"))
            .add_run(
                Run::new(format!(" TOC \\o \"1-{}\" \\h \\z \\u ", config.depth)).with_instr_text(),
            )
            .add_run(Run::new("").with_field_char("separate"));
        elements.push(DocElement::Paragraph(Box::new(toc_field_begin)));

        // 3. Static placeholder entries (Word updates these when field is updated)
        // Each entry has: text, tab, and PAGEREF field for page number
        for entry in self.entries.iter().filter(|e| e.level <= config.depth) {
            let style = format!("TOC{}", entry.level);

            // Create TOC entry with tab and page reference
            let toc_para = Paragraph::with_style(&style)
                .spacing(0, 0)
                .line_spacing(240, "auto")
                .add_run(Run::new(&entry.text))
                .add_run(Run::new("").with_tab())
                .add_run(Run::new("").with_field_char("begin"))
                .add_run(Run::new(format!(" PAGEREF {} \\h ", entry.bookmark_id)).with_instr_text())
                .add_run(Run::new("").with_field_char("separate"))
                .add_run(Run::new("1")) // Placeholder page number
                .add_run(Run::new("").with_field_char("end"));

            elements.push(DocElement::Paragraph(Box::new(toc_para)));
        }

        // 4. TOC Field end
        let toc_field_end = Paragraph::new()
            .spacing(0, 0)
            .line_spacing(240, "auto")
            .add_run(Run::new("").with_field_char("end"));
        elements.push(DocElement::Paragraph(Box::new(toc_field_end)));

        // 5. Section break after TOC
        let section_break = Paragraph::new()
            .spacing(0, 0)
            .line_spacing(240, "auto")
            .section_break("nextPage")
            .suppress_header_footer();
        elements.push(DocElement::Paragraph(Box::new(section_break)));

        elements
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_toc_builder_add_heading() {
        let mut builder = TocBuilder::new();
        let id1 = builder.add_heading(1, "Introduction", None);
        let id2 = builder.add_heading(2, "Getting Started", None);
        let id3 = builder.add_heading(1, "Conclusion", Some("conclusion"));

        assert!(id1.starts_with("_Toc"));
        assert!(id1.contains("Introduction"));
        assert!(id2.contains("Getting_Started"));
        assert_eq!(id3, "conclusion"); // Explicit ID used as-is
        assert_eq!(builder.entries().len(), 3);
    }

    #[test]
    fn test_toc_builder_generate_toc() {
        let mut builder = TocBuilder::new();
        builder.add_heading(1, "Chapter 1", None);
        builder.add_heading(2, "Section 1.1", None);
        builder.add_heading(3, "Subsection 1.1.1", None);
        builder.add_heading(4, "Deep heading", None); // Should be filtered out with depth=2

        let config = TocConfig::default(); // depth = 2
        let elements = builder.generate_toc(&config);

        // Should have: title + field begin + 2 entries (h1, h2) + field end + section break
        // = 1 + 1 + 2 + 1 + 1 = 6 elements
        assert_eq!(elements.len(), 6);
    }

    #[test]
    fn test_toc_config_default() {
        let config = TocConfig::default();
        assert!(config.enabled);
        assert_eq!(config.depth, 2); // Changed from 3 to 2
        assert_eq!(config.title, "Table of Contents");
        assert!(config.after_cover); // New field
    }

    #[test]
    fn test_toc_disabled() {
        let builder = TocBuilder::new();
        let config = TocConfig {
            enabled: false,
            ..Default::default()
        };
        let elements = builder.generate_toc(&config);
        assert!(elements.is_empty());
    }

    #[test]
    fn test_bookmark_id_sanitization() {
        let mut builder = TocBuilder::new();
        let id = builder.add_heading(1, "Hello World! @#$%", None);

        // Should only contain alphanumeric and underscores
        assert!(id.chars().all(|c| c.is_alphanumeric() || c == '_'));
        assert!(id.contains("Hello_World"));
    }

    #[test]
    fn test_toc_builder_is_empty() {
        let builder = TocBuilder::new();
        assert!(builder.is_empty());

        let mut builder = TocBuilder::new();
        builder.add_heading(1, "Test", None);
        assert!(!builder.is_empty());
    }

    #[test]
    fn test_toc_depth_filtering() {
        let mut builder = TocBuilder::new();
        builder.add_heading(1, "H1", None);
        builder.add_heading(2, "H2", None);
        builder.add_heading(3, "H3", None);
        builder.add_heading(4, "H4", None);
        builder.add_heading(5, "H5", None);

        // Test depth = 2
        let config = TocConfig {
            enabled: true,
            depth: 2,
            title: "TOC".to_string(),
            after_cover: true,
        };
        let elements = builder.generate_toc(&config);

        // Should have: title + field begin + 2 entries (H1 and H2) + field end + section break
        // = 1 + 1 + 2 + 1 + 1 = 6 elements
        assert_eq!(elements.len(), 6);
    }

    #[test]
    fn test_toc_empty_entries() {
        let builder = TocBuilder::new();
        let config = TocConfig::default();
        let elements = builder.generate_toc(&config);

        // Should be empty when no headings added
        assert!(elements.is_empty());
    }

    #[test]
    fn test_toc_entry_structure() {
        let mut builder = TocBuilder::new();
        builder.add_heading(2, "Test Heading", None);

        let entries = builder.entries();
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].text, "Test Heading");
        assert_eq!(entries[0].level, 2);
        assert!(entries[0].bookmark_id.starts_with("_Toc"));
    }

    #[test]
    fn test_toc_custom_title() {
        let mut builder = TocBuilder::new();
        builder.add_heading(1, "Chapter 1", None);

        let config = TocConfig {
            enabled: true,
            depth: 2,
            title: "Contents".to_string(),
            after_cover: true,
        };
        let elements = builder.generate_toc(&config);

        // Should have title paragraph with custom title
        assert_eq!(elements.len(), 5); // title + field begin + entry + field end + section break
        match &elements[0] {
            DocElement::Paragraph(p) => {
                assert_eq!(p.style_id, Some("TOCHeading".to_string()));
                assert!(!p.children.is_empty());
                match &p.children[0] {
                    crate::docx::ooxml::ParagraphChild::Run(run) => {
                        assert_eq!(run.text, "Contents");
                    }
                    _ => panic!("Expected Run child"),
                }
            }
            _ => panic!("Expected Paragraph element"),
        }
    }

    #[test]
    fn test_toc_multiple_headings_same_level() {
        let mut builder = TocBuilder::new();
        builder.add_heading(1, "Chapter 1", None);
        builder.add_heading(2, "Section 1.1", None);
        builder.add_heading(2, "Section 1.2", None);
        builder.add_heading(1, "Chapter 2", None);

        let config = TocConfig::default();
        let elements = builder.generate_toc(&config);

        // Should have: title + field begin + 4 entries + field end + section break
        // = 1 + 1 + 4 + 1 + 1 = 8 elements
        assert_eq!(elements.len(), 8);
    }

    #[test]
    fn test_toc_bookmark_id_uniqueness() {
        let mut builder = TocBuilder::new();
        let id1 = builder.add_heading(1, "Introduction", None);
        let id2 = builder.add_heading(1, "Introduction", None);

        // Even with same text, IDs should be unique due to counter
        assert_ne!(id1, id2);
    }

    #[test]
    fn test_toc_bookmark_id_max_length() {
        let mut builder = TocBuilder::new();
        let long_text =
            "This is a very long heading text that should be truncated to 40 characters";
        let id = builder.add_heading(1, long_text, None);

        // ID should be truncated (excluding the "_TocN_" prefix)
        let text_part = id.split('_').last().unwrap_or("");
        assert!(text_part.len() <= 40);
    }
}
