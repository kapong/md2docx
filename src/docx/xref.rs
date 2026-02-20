//! Cross-reference context for tracking anchors and resolving references

use crate::parser::RefType;
use std::collections::HashMap;

/// Information about a registered anchor/bookmark
#[derive(Debug, Clone)]
pub(crate) struct AnchorInfo {
    #[allow(dead_code)]
    pub id: String, // User-defined ID (e.g., "intro", "arch")
    pub bookmark_name: String,  // OOXML bookmark name (e.g., "_Ref_intro")
    pub ref_type: RefType,      // Chapter, Figure, Table, etc.
    pub display_text: String,   // Text shown when referenced (e.g., "Introduction")
    pub number: Option<String>, // Numbering like "1.2" for figures
}

/// Context for tracking cross-references during document build
#[derive(Debug, Default)]
pub(crate) struct CrossRefContext {
    /// Map from anchor ID to anchor info
    anchors: HashMap<String, AnchorInfo>,
    /// Counter for generating unique bookmark IDs
    next_bookmark_id: u32,
    /// Counters for numbering
    chapter_num: u32,
    figure_num: u32,
    table_num: u32,
    equation_num: u32,
}

impl CrossRefContext {
    pub fn new() -> Self {
        Self::default()
    }

    /// Register a heading anchor
    /// Called when processing Block::Heading with an id
    pub fn register_heading(&mut self, id: &str, level: u8, text: &str) -> String {
        self.next_bookmark_id += 1;
        let bookmark_name = format!("_Ref_{}", sanitize_bookmark_name(id));

        // Determine ref type and numbering based on level
        let (ref_type, number) = if level == 1 {
            self.chapter_num += 1;
            self.figure_num = 0; // Reset per-chapter counters
            self.table_num = 0;
            self.equation_num = 0;
            (RefType::Chapter, Some(self.chapter_num.to_string()))
        } else {
            (RefType::Section, None)
        };

        self.anchors.insert(
            id.to_string(),
            AnchorInfo {
                id: id.to_string(),
                bookmark_name: bookmark_name.clone(),
                ref_type,
                display_text: text.to_string(),
                number,
            },
        );

        bookmark_name
    }

    /// Register a figure anchor
    pub fn register_figure(&mut self, id: &str, alt_text: &str) -> String {
        self.next_bookmark_id += 1;
        self.figure_num += 1;

        let bookmark_name = format!("_Ref_{}", sanitize_bookmark_name(id));
        let number = if self.chapter_num > 0 {
            format!("{}.{}", self.chapter_num, self.figure_num)
        } else {
            self.figure_num.to_string()
        };

        self.anchors.insert(
            id.to_string(),
            AnchorInfo {
                id: id.to_string(),
                bookmark_name: bookmark_name.clone(),
                ref_type: RefType::Figure,
                display_text: alt_text.to_string(),
                number: Some(number),
            },
        );

        bookmark_name
    }

    /// Register a table anchor
    pub fn register_table(&mut self, id: &str, caption: &str) -> String {
        self.next_bookmark_id += 1;
        self.table_num += 1;

        let bookmark_name = format!("_Ref_{}", sanitize_bookmark_name(id));
        let number = if self.chapter_num > 0 {
            format!("{}.{}", self.chapter_num, self.table_num)
        } else {
            self.table_num.to_string()
        };

        self.anchors.insert(
            id.to_string(),
            AnchorInfo {
                id: id.to_string(),
                bookmark_name: bookmark_name.clone(),
                ref_type: RefType::Table,
                display_text: caption.to_string(),
                number: Some(number),
            },
        );

        bookmark_name
    }

    /// Register an equation anchor
    pub fn register_equation(&mut self, id: &str) -> String {
        self.next_bookmark_id += 1;
        self.equation_num += 1;

        let bookmark_name = format!("_Ref_{}", sanitize_bookmark_name(id));
        let number = if self.chapter_num > 0 {
            format!("{}.{}", self.chapter_num, self.equation_num)
        } else {
            self.equation_num.to_string()
        };

        self.anchors.insert(
            id.to_string(),
            AnchorInfo {
                id: id.to_string(),
                bookmark_name: bookmark_name.clone(),
                ref_type: RefType::Equation,
                display_text: format!("Equation {}", number),
                number: Some(number),
            },
        );

        bookmark_name
    }

    /// Get current equation number (for display equations without an explicit id)
    pub fn next_equation_number(&mut self) -> String {
        self.equation_num += 1;
        if self.chapter_num > 0 {
            format!("{}.{}", self.chapter_num, self.equation_num)
        } else {
            self.equation_num.to_string()
        }
    }

    /// Register a generic anchor (for future extensibility)
    #[allow(dead_code)]
    pub fn register_anchor(&mut self, id: &str, ref_type: RefType, text: &str) -> String {
        self.next_bookmark_id += 1;
        let bookmark_name = format!("_Ref_{}", sanitize_bookmark_name(id));

        self.anchors.insert(
            id.to_string(),
            AnchorInfo {
                id: id.to_string(),
                bookmark_name: bookmark_name.clone(),
                ref_type,
                display_text: text.to_string(),
                number: None,
            },
        );

        bookmark_name
    }

    /// Resolve a cross-reference by target ID
    /// Returns the anchor info if found
    pub fn resolve(&self, target: &str) -> Option<&AnchorInfo> {
        self.anchors.get(target)
    }

    /// Get display text for a reference
    /// Returns formatted text like "Figure 1.2" or just the title
    #[allow(dead_code)]
    pub fn get_display_text(&self, target: &str, _ref_type: RefType) -> String {
        self.get_localized_display_text(target, crate::docx::ooxml::Language::English)
    }

    /// Get localized display text for a reference
    pub fn get_localized_display_text(
        &self,
        target: &str,
        lang: crate::docx::ooxml::Language,
    ) -> String {
        if let Some(anchor) = self.anchors.get(target) {
            match anchor.ref_type {
                RefType::Figure => {
                    if let Some(num) = &anchor.number {
                        format!("{} {}", lang.figure_caption_prefix(), num)
                    } else {
                        anchor.display_text.clone()
                    }
                }
                RefType::Table => {
                    if let Some(num) = &anchor.number {
                        format!("{} {}", lang.table_caption_prefix(), num)
                    } else {
                        anchor.display_text.clone()
                    }
                }
                RefType::Chapter => {
                    if let Some(num) = &anchor.number {
                        match lang {
                            crate::docx::ooxml::Language::Thai => format!("บทที่ {}", num),
                            _ => format!("Chapter {}", num),
                        }
                    } else {
                        anchor.display_text.clone()
                    }
                }
                RefType::Section => anchor.display_text.clone(),
                RefType::Equation => {
                    if let Some(num) = &anchor.number {
                        num.clone()
                    } else {
                        anchor.display_text.clone()
                    }
                }
                RefType::Appendix => {
                    if let Some(num) = &anchor.number {
                        match lang {
                            crate::docx::ooxml::Language::Thai => format!("ภาคผนวก {}", num),
                            _ => format!("Appendix {}", num),
                        }
                    } else {
                        anchor.display_text.clone()
                    }
                }
                _ => anchor.display_text.clone(),
            }
        } else {
            // Reference not found - return placeholder
            format!("[{}]", target)
        }
    }

    /// Check if an anchor exists
    #[allow(dead_code)]
    pub fn has_anchor(&self, id: &str) -> bool {
        self.anchors.contains_key(id)
    }

    /// Get all registered anchors (for debugging/testing)
    #[allow(dead_code)]
    pub fn anchors(&self) -> &HashMap<String, AnchorInfo> {
        &self.anchors
    }
}

/// Sanitize a string for use as a bookmark name
/// Keeps only alphanumeric and underscores
fn sanitize_bookmark_name(s: &str) -> String {
    s.chars()
        .filter(|c| c.is_alphanumeric() || *c == '_')
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_register_heading() {
        let mut ctx = CrossRefContext::new();
        let bookmark = ctx.register_heading("intro", 1, "Introduction");

        assert!(bookmark.starts_with("_Ref_"));
        assert!(ctx.has_anchor("intro"));

        let anchor = ctx.resolve("intro").unwrap();
        assert_eq!(anchor.display_text, "Introduction");
        assert_eq!(anchor.ref_type, RefType::Chapter);
        assert_eq!(anchor.number, Some("1".to_string()));
    }

    #[test]
    fn test_register_figure() {
        let mut ctx = CrossRefContext::new();
        ctx.register_heading("ch1", 1, "Chapter 1"); // Set chapter context
        let bookmark = ctx.register_figure("arch", "System Architecture");

        assert!(bookmark.contains("arch"));
        let anchor = ctx.resolve("arch").unwrap();
        assert_eq!(anchor.ref_type, RefType::Figure);
        assert_eq!(anchor.number, Some("1.1".to_string()));
    }

    #[test]
    fn test_register_table() {
        let mut ctx = CrossRefContext::new();
        ctx.register_heading("ch1", 1, "Chapter 1");
        let _bookmark = ctx.register_table("users", "User Data");

        let anchor = ctx.resolve("users").unwrap();
        assert_eq!(anchor.ref_type, RefType::Table);
        assert_eq!(anchor.number, Some("1.1".to_string()));
    }

    #[test]
    fn test_get_display_text() {
        let mut ctx = CrossRefContext::new();
        ctx.register_heading("ch1", 1, "Getting Started");
        ctx.register_figure("diagram", "Overview Diagram");

        assert_eq!(ctx.get_display_text("ch1", RefType::Chapter), "Chapter 1");
        assert_eq!(
            ctx.get_display_text("diagram", RefType::Figure),
            "Figure 1.1"
        );
        assert_eq!(
            ctx.get_display_text("unknown", RefType::Unknown),
            "[unknown]"
        );
    }

    #[test]
    fn test_get_localized_display_text() {
        let mut ctx = CrossRefContext::new();
        ctx.register_heading("ch1", 1, "Getting Started");
        ctx.register_table("users", "User List");
        ctx.register_figure("diagram", "Overview Diagram");

        use crate::docx::ooxml::Language;

        // English
        assert_eq!(
            ctx.get_localized_display_text("ch1", Language::English),
            "Chapter 1"
        );
        assert_eq!(
            ctx.get_localized_display_text("users", Language::English),
            "Table 1.1"
        );
        assert_eq!(
            ctx.get_localized_display_text("diagram", Language::English),
            "Figure 1.1"
        );

        // Thai
        assert_eq!(
            ctx.get_localized_display_text("ch1", Language::Thai),
            "บทที่ 1"
        );
        assert_eq!(
            ctx.get_localized_display_text("users", Language::Thai),
            "ตารางที่ 1.1"
        );
        assert_eq!(
            ctx.get_localized_display_text("diagram", Language::Thai),
            "รูปที่ 1.1"
        );
    }

    #[test]
    fn test_chapter_resets_counters() {
        let mut ctx = CrossRefContext::new();
        ctx.register_heading("ch1", 1, "Chapter 1");
        ctx.register_figure("fig1", "Figure in Ch1");

        ctx.register_heading("ch2", 1, "Chapter 2");
        ctx.register_figure("fig2", "Figure in Ch2");

        let fig1 = ctx.resolve("fig1").unwrap();
        let fig2 = ctx.resolve("fig2").unwrap();

        assert_eq!(fig1.number, Some("1.1".to_string()));
        assert_eq!(fig2.number, Some("2.1".to_string()));
    }

    #[test]
    fn test_sanitize_bookmark_name() {
        assert_eq!(sanitize_bookmark_name("hello-world"), "helloworld");
        assert_eq!(sanitize_bookmark_name("fig:arch"), "figarch");
        assert_eq!(sanitize_bookmark_name("test_123"), "test_123");
    }
}
