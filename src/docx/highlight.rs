//! Syntax highlighting for code blocks using syntect.
//!
//! Converts source code into a sequence of colored text runs
//! that can be rendered in DOCX.

use once_cell::sync::Lazy;
use syntect::highlighting::{Color, ThemeSet};
use syntect::parsing::SyntaxSet;

/// Pre-loaded syntax and theme sets (loaded once, reused for all code blocks).
static SYNTAX_SET: Lazy<SyntaxSet> = Lazy::new(SyntaxSet::load_defaults_newlines);
static THEME_SET: Lazy<ThemeSet> = Lazy::new(ThemeSet::load_defaults);

/// A highlighted token: (text, optional hex color without '#').
pub type HighlightedToken = (String, Option<String>);

/// A highlighted line is a list of tokens.
pub type HighlightedLine = Vec<HighlightedToken>;

/// Convert a syntect `Color` to a hex string (without `#`).
fn color_to_hex(c: Color) -> String {
    format!("{:02X}{:02X}{:02X}", c.r, c.g, c.b)
}

/// Highlight source code and return one `HighlightedLine` per line.
///
/// If the language is not recognised, or is `None`, the code is returned
/// as plain (uncolored) text.
pub fn highlight_code(code: &str, lang: Option<&str>) -> Vec<HighlightedLine> {
    // Try to find a syntax definition for the language
    let syntax = lang
        .and_then(|l| {
            // Skip mermaid – it's a diagram language, not highlighted code
            if l == "mermaid" {
                return None;
            }
            SYNTAX_SET
                .find_syntax_by_token(l)
                .or_else(|| SYNTAX_SET.find_syntax_by_extension(l))
        })
        .unwrap_or_else(|| SYNTAX_SET.find_syntax_plain_text());

    let theme = &THEME_SET.themes["InspiredGitHub"];

    // Default foreground from the theme (used to skip emitting color for "normal" text)
    let default_fg = theme
        .settings
        .foreground
        .unwrap_or(Color {
            r: 0,
            g: 0,
            b: 0,
            a: 255,
        });

    let mut highlighter = syntect::easy::HighlightLines::new(syntax, theme);
    let mut result: Vec<HighlightedLine> = Vec::new();

    for line in syntect::util::LinesWithEndings::from(code) {
        let ranges = highlighter
            .highlight_line(line, &SYNTAX_SET)
            .unwrap_or_default();

        let mut tokens = Vec::new();
        for (style, text) in ranges {
            // Strip the trailing newline – we handle lines via paragraphs
            let text = text.trim_end_matches('\n').trim_end_matches('\r');
            if text.is_empty() {
                continue;
            }

            let color = if style.foreground == default_fg {
                None // use default font color
            } else {
                Some(color_to_hex(style.foreground))
            };

            tokens.push((text.to_string(), color));
        }
        result.push(tokens);
    }

    // Ensure at least one (empty) line if the code is empty
    if result.is_empty() {
        result.push(Vec::new());
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_highlight_plain_text() {
        let lines = highlight_code("hello world", None);
        assert_eq!(lines.len(), 1);
        // Plain text should have tokens
        assert!(!lines[0].is_empty());
        // Concatenated text matches input (trimmed)
        let joined: String = lines[0].iter().map(|(t, _)| t.as_str()).collect();
        assert_eq!(joined, "hello world");
    }

    #[test]
    fn test_highlight_rust() {
        let code = "fn main() {\n    println!(\"hello\");\n}\n";
        let lines = highlight_code(code, Some("rust"));
        assert!(lines.len() >= 3);
        // First token of first line should be `fn` keyword, likely with a color
        let first_text: String = lines[0].iter().map(|(t, _)| t.as_str()).collect();
        assert!(first_text.contains("fn"));
    }

    #[test]
    fn test_highlight_unknown_lang() {
        let lines = highlight_code("some code", Some("unknown_lang_xyz"));
        assert_eq!(lines.len(), 1);
        let joined: String = lines[0].iter().map(|(t, _)| t.as_str()).collect();
        assert_eq!(joined, "some code");
    }

    #[test]
    fn test_highlight_empty() {
        let lines = highlight_code("", None);
        assert_eq!(lines.len(), 1);
        assert!(lines[0].is_empty());
    }

    #[test]
    fn test_highlight_python() {
        let code = "def add(a, b):\n    return a + b\n";
        let lines = highlight_code(code, Some("python"));
        assert!(lines.len() >= 2);
        // `def` keyword should be highlighted with a color
        let has_color = lines[0].iter().any(|(_, c)| c.is_some());
        assert!(has_color, "Python keyword 'def' should be syntax-highlighted");
    }
}
