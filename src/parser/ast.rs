//! Abstract Syntax Tree for parsed Markdown
//!
//! This module defines the intermediate representation between
//! Markdown parsing and DOCX generation.

use std::collections::HashMap;

/// A complete parsed document
#[derive(Debug, Clone, Default)]
pub struct ParsedDocument {
    /// YAML frontmatter metadata
    pub frontmatter: Option<Frontmatter>,
    /// Document content blocks
    pub blocks: Vec<Block>,
    /// Map of footnote label (e.g. "1") to content blocks
    pub footnotes: HashMap<String, Vec<Block>>,
}

/// YAML frontmatter metadata
#[derive(Debug, Clone, Default)]
pub struct Frontmatter {
    pub title: Option<String>,
    pub title_th: Option<String>,
    pub skip_toc: bool,
    pub skip_numbering: bool,
    pub page_break_before: bool,
    pub header_override: Option<String>,
    pub language: Option<String>,
    /// Additional custom fields
    pub extra: HashMap<String, String>,
}

/// Block-level elements
#[derive(Debug, Clone)]
pub enum Block {
    /// Heading with level (1-6), content, and optional anchor ID
    Heading {
        level: u8,
        content: Vec<Inline>,
        id: Option<String>,
    },

    /// Regular paragraph
    Paragraph(Vec<Inline>),

    /// Code block with optional language and metadata
    CodeBlock {
        lang: Option<String>,
        content: String,
        filename: Option<String>,
        highlight_lines: Vec<u32>,
        show_line_numbers: bool,
    },

    /// Block quote (can contain nested blocks)
    BlockQuote(Vec<Block>),

    /// List (ordered or unordered)
    List {
        ordered: bool,
        start: Option<u32>, // Starting number for ordered lists
        items: Vec<ListItem>,
    },

    /// Table
    Table {
        headers: Vec<TableCell>,
        alignments: Vec<Alignment>,
        rows: Vec<Vec<TableCell>>,
        caption: Option<String>,
        id: Option<String>,
    },

    /// Image (block-level, becomes figure with caption)
    Image {
        alt: String,
        src: String,
        title: Option<String>,
        width: Option<String>,
        id: Option<String>, // For cross-references
    },

    /// Horizontal rule / thematic break
    ThematicBreak,

    /// Mermaid diagram
    Mermaid { content: String, id: Option<String> },

    /// Raw HTML (preserved but may not render in DOCX)
    Html(String),

    /// Math block (display equation): $$...$$
    MathBlock { content: String },

    /// Include directive: {!include:path.md}
    Include {
        path: String,
        resolved: Option<Vec<Block>>, // Filled after resolution
    },

    /// Code include: {!code:src/main.rs:10-25}
    CodeInclude {
        path: String,
        start_line: Option<u32>,
        end_line: Option<u32>,
        lang: Option<String>,
    },
}

/// List item (can contain nested blocks)
#[derive(Debug, Clone)]
pub struct ListItem {
    pub content: Vec<Block>,
    pub checked: Option<bool>, // For task lists: Some(true), Some(false), or None
}

/// Table cell
#[derive(Debug, Clone)]
pub struct TableCell {
    pub content: Vec<Inline>,
    pub is_header: bool,
}

/// Table column alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Alignment {
    Left,
    Center,
    Right,
    #[default]
    None,
}

/// Inline elements
#[derive(Debug, Clone, PartialEq)]
pub enum Inline {
    /// Plain text
    Text(String),

    /// Bold/strong text
    Bold(Vec<Inline>),

    /// Italic/emphasis text
    Italic(Vec<Inline>),

    /// Bold + Italic
    BoldItalic(Vec<Inline>),

    /// Inline code
    Code(String),

    /// Strikethrough text
    Strikethrough(Vec<Inline>),

    /// Hyperlink
    Link {
        text: Vec<Inline>,
        url: String,
        title: Option<String>,
    },

    /// Inline image
    Image {
        alt: String,
        src: String,
        title: Option<String>,
    },

    /// Footnote reference
    FootnoteRef(String),

    /// Cross-reference: {ref:ch02} or {ref:fig:diagram}
    CrossRef { target: String, ref_type: RefType },

    /// Soft break (single newline in source)
    SoftBreak,

    /// Hard break (two spaces + newline or <br>)
    HardBreak,

    /// Raw HTML inline
    Html(String),

    /// Index marker: {index:term}
    IndexMarker(String),

    /// Inline math: $...$
    InlineMath(String),

    /// Display math (inline context): $$...$$
    DisplayMath(String),
}

/// Extract plain text from inline elements
pub fn extract_inline_text(inlines: &[Inline]) -> String {
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
            Inline::InlineMath(s) | Inline::DisplayMath(s) => s.clone(),
        })
        .collect::<Vec<_>>()
        .join("")
}

/// Type of cross-reference
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RefType {
    Chapter,
    Section,
    Figure,
    Table,
    Appendix,
    Footnote,
    Unknown,
}

impl RefType {
    pub fn from_prefix(s: &str) -> Self {
        match s.to_lowercase().as_str() {
            "ch" | "chapter" => RefType::Chapter,
            "sec" | "section" => RefType::Section,
            "fig" | "figure" => RefType::Figure,
            "tbl" | "table" => RefType::Table,
            "ap" | "appendix" => RefType::Appendix,
            "fn" | "footnote" => RefType::Footnote,
            _ => RefType::Unknown,
        }
    }
}

// Implement helper constructors
impl Block {
    pub fn heading(level: u8, text: &str) -> Self {
        Block::Heading {
            level,
            content: vec![Inline::Text(text.to_string())],
            id: None,
        }
    }

    pub fn paragraph(text: &str) -> Self {
        Block::Paragraph(vec![Inline::Text(text.to_string())])
    }

    pub fn code_block(content: &str, lang: Option<&str>) -> Self {
        Block::CodeBlock {
            lang: lang.map(|s| s.to_string()),
            content: content.to_string(),
            filename: None,
            highlight_lines: Vec::new(),
            show_line_numbers: false,
        }
    }
}

impl Inline {
    pub fn text(s: &str) -> Self {
        Inline::Text(s.to_string())
    }

    pub fn bold(content: Vec<Inline>) -> Self {
        Inline::Bold(content)
    }

    pub fn italic(content: Vec<Inline>) -> Self {
        Inline::Italic(content)
    }

    pub fn code(s: &str) -> Self {
        Inline::Code(s.to_string())
    }

    pub fn link(text: &str, url: &str) -> Self {
        Inline::Link {
            text: vec![Inline::Text(text.to_string())],
            url: url.to_string(),
            title: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_block_heading() {
        let h = Block::heading(1, "Test");
        match h {
            Block::Heading { level, content, id } => {
                assert_eq!(level, 1);
                assert!(id.is_none());
                assert_eq!(content.len(), 1);
            }
            _ => panic!("Expected Heading"),
        }
    }

    #[test]
    fn test_inline_text() {
        let t = Inline::text("Hello");
        match t {
            Inline::Text(s) => assert_eq!(s, "Hello"),
            _ => panic!("Expected Text"),
        }
    }

    #[test]
    fn test_ref_type_from_prefix() {
        assert_eq!(RefType::from_prefix("fig"), RefType::Figure);
        assert_eq!(RefType::from_prefix("ch"), RefType::Chapter);
        assert_eq!(RefType::from_prefix("unknown"), RefType::Unknown);
    }
}
