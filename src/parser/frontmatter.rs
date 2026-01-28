//! Frontmatter parser
//!
//! Extracts YAML metadata from the beginning of markdown files.
//! Frontmatter is delimited by `---` markers:
//!
//! ```markdown
//! ---
//! title: "My Document"
//! skip_toc: false
//! ---
//!
//! # Content here
//! ```

use crate::parser::ast::{Frontmatter, ParsedDocument};

/// Parse frontmatter from markdown content
///
/// Returns (frontmatter, remaining_markdown_content)
///
/// # Rules
/// - Frontmatter must start at the very beginning of the file
/// - Opening `---` must be on line 1
/// - Closing `---` must be on its own line
/// - YAML between the markers is parsed into `Frontmatter`
/// - All known fields map to struct fields, unknown keys go to `extra` HashMap
/// - If no frontmatter found, return `(None, input)` - the full input as remaining
///
/// # Examples
///
/// ```rust
/// use md2docx::parser::parse_frontmatter;
///
/// let md = r#"---
/// title: "Getting Started"
/// skip_toc: false
/// ---
///
/// # Chapter 1
/// "#;
///
/// let (frontmatter, _content) = parse_frontmatter(md);
/// assert!(frontmatter.is_some());
/// let fm = frontmatter.unwrap();
/// assert_eq!(fm.title, Some("Getting Started".to_string()));
/// assert_eq!(fm.skip_toc, false);
/// // Content contains the remaining markdown after frontmatter
/// ```
pub fn parse_frontmatter(input: &str) -> (Option<Frontmatter>, &str) {
    // Check if input starts with ---
    if !input.starts_with("---") {
        return (None, input);
    }

    // Find the closing ---
    let lines: Vec<&str> = input.lines().collect();

    // Must have at least 2 lines (opening and closing)
    if lines.len() < 2 {
        return (None, input);
    }

    // Find closing marker (must be on its own line)
    let mut closing_line = None;
    for (i, line) in lines.iter().enumerate().skip(1) {
        if line.trim() == "---" {
            closing_line = Some(i);
            break;
        }
    }

    let closing_idx = match closing_line {
        Some(idx) => idx,
        None => return (None, input), // No closing marker found
    };

    // Extract YAML content (between opening and closing ---)
    let yaml_lines = &lines[1..closing_idx];
    let yaml_content = yaml_lines.join("\n");

    // Parse YAML into Frontmatter
    let frontmatter = parse_yaml_frontmatter(&yaml_content);

    // Calculate remaining content (everything after closing ---)
    // Find the byte position of the closing --- line
    let mut byte_pos = 0;
    for (i, line) in lines.iter().enumerate() {
        if i == closing_idx {
            // Position after the closing line (including its newline)
            byte_pos += line.len();
            // Add newline character(s) - either \n or \r\n
            if byte_pos < input.len() && input.as_bytes()[byte_pos] == b'\r' {
                byte_pos += 1; // Skip \r
            }
            if byte_pos < input.len() && input.as_bytes()[byte_pos] == b'\n' {
                byte_pos += 1; // Skip \n
            }
            break;
        }
        byte_pos += line.len();
        // Add newline character(s)
        if byte_pos < input.len() && input.as_bytes()[byte_pos] == b'\r' {
            byte_pos += 1;
        }
        if byte_pos < input.len() && input.as_bytes()[byte_pos] == b'\n' {
            byte_pos += 1;
        }
    }

    let remaining = if byte_pos < input.len() {
        &input[byte_pos..]
    } else {
        ""
    };

    (frontmatter, remaining)
}

/// Parse YAML content into Frontmatter struct
///
/// Handles simple key: value pairs without full YAML parser
fn parse_yaml_frontmatter(yaml: &str) -> Option<Frontmatter> {
    let mut frontmatter = Frontmatter::default();

    for line in yaml.lines() {
        let line = line.trim();

        // Skip empty lines and comments
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        // Parse key: value
        if let Some(colon_pos) = line.find(':') {
            let key = line[..colon_pos].trim();
            let value = line[colon_pos + 1..].trim();

            // Skip if value is empty (might be a nested structure)
            if value.is_empty() {
                continue;
            }

            // Parse value (handle quoted and unquoted strings)
            let parsed_value = parse_yaml_value(value);

            // Map known fields
            match key {
                "title" => frontmatter.title = parsed_value,
                "title_th" => frontmatter.title_th = parsed_value,
                "skip_toc" => frontmatter.skip_toc = parse_bool(value),
                "skip_numbering" => frontmatter.skip_numbering = parse_bool(value),
                "page_break_before" => frontmatter.page_break_before = parse_bool(value),
                "header_override" => frontmatter.header_override = parsed_value,
                "language" | "lang" => frontmatter.language = parsed_value,
                _ => {
                    // Unknown keys go to extra HashMap
                    if let Some(val) = parsed_value {
                        frontmatter.extra.insert(key.to_string(), val);
                    }
                }
            }
        }
    }

    Some(frontmatter)
}

/// Parse a YAML value (handles quoted and unquoted strings)
fn parse_yaml_value(value: &str) -> Option<String> {
    let value = value.trim();

    if value.is_empty() {
        return None;
    }

    // Handle quoted strings
    if (value.starts_with('"') && value.ends_with('"'))
        || (value.starts_with('\'') && value.ends_with('\''))
    {
        // Remove quotes
        let unquoted = &value[1..value.len() - 1];
        Some(unquoted.to_string())
    } else {
        // Handle unquoted strings (take everything after colon)
        Some(value.to_string())
    }
}

/// Parse a boolean value from YAML
fn parse_bool(value: &str) -> bool {
    let value = value.trim().to_lowercase();

    match value.as_str() {
        "true" | "yes" | "1" => true,
        "false" | "no" | "0" => false,
        // Handle quoted booleans
        s if s == "\"true\"" || s == "'true'" => true,
        s if s == "\"false\"" || s == "'false'" => false,
        _ => false, // Default to false for unknown values
    }
}

/// Parse markdown with frontmatter extraction
///
/// This calls parse_frontmatter first, then parse_markdown on the remaining content
///
/// # Examples
///
/// ```rust
/// use md2docx::parser::parse_markdown_with_frontmatter;
///
/// let md = r#"---
/// title: "Getting Started"
/// skip_toc: false
/// ---
///
/// # Chapter 1
///
/// Content here.
/// "#;
///
/// let doc = parse_markdown_with_frontmatter(md);
/// assert!(doc.frontmatter.is_some());
/// assert!(!doc.blocks.is_empty());
/// ```
pub fn parse_markdown_with_frontmatter(input: &str) -> ParsedDocument {
    let (frontmatter, content) = parse_frontmatter(input);
    let mut doc = crate::parser::markdown::parse_markdown(content);
    doc.frontmatter = frontmatter;
    doc
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_frontmatter_basic() {
        let md = r#"---
title: "Getting Started"
skip_toc: false
---

# Chapter 1

Content here.
"#;

        let (frontmatter, content) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Getting Started".to_string()));
        assert_eq!(fm.skip_toc, false);
        assert!(content.starts_with("\n# Chapter 1"));
    }

    #[test]
    fn test_parse_frontmatter_with_thai() {
        let md = r#"---
title: "Getting Started"
title_th: "เริ่มต้นใช้งาน"
skip_toc: false
---

# Chapter 1
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Getting Started".to_string()));
        assert_eq!(fm.title_th, Some("เริ่มต้นใช้งาน".to_string()));
    }

    #[test]
    fn test_parse_frontmatter_all_fields() {
        let md = r#"---
title: "Custom Title"
title_th: "หัวข้อ"
skip_toc: true
skip_numbering: false
page_break_before: true
header_override: "Special Section"
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Custom Title".to_string()));
        assert_eq!(fm.title_th, Some("หัวข้อ".to_string()));
        assert_eq!(fm.skip_toc, true);
        assert_eq!(fm.skip_numbering, false);
        assert_eq!(fm.page_break_before, true);
        assert_eq!(fm.header_override, Some("Special Section".to_string()));
    }

    #[test]
    fn test_parse_frontmatter_with_extra_fields() {
        let md = r#"---
title: "Test"
custom_key: "custom_value"
another_key: another_value
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Test".to_string()));
        assert_eq!(
            fm.extra.get("custom_key"),
            Some(&"custom_value".to_string())
        );
        assert_eq!(
            fm.extra.get("another_key"),
            Some(&"another_value".to_string())
        );
    }

    #[test]
    fn test_parse_frontmatter_no_frontmatter() {
        let md = "# Chapter 1\n\nContent here.";

        let (frontmatter, content) = parse_frontmatter(md);
        assert!(frontmatter.is_none());
        assert_eq!(content, md);
    }

    #[test]
    fn test_parse_frontmatter_empty() {
        let md = r#"---
---

Content
"#;

        let (frontmatter, content) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, None);
        assert_eq!(fm.skip_toc, false);
        assert_eq!(fm.skip_numbering, false);
        assert_eq!(fm.page_break_before, false);
        assert!(content.starts_with("\nContent"));
    }

    #[test]
    fn test_parse_frontmatter_no_closing_marker() {
        let md = r#"---
title: "Test"

# Content
"#;

        let (frontmatter, content) = parse_frontmatter(md);
        assert!(frontmatter.is_none());
        assert_eq!(content, md);
    }

    #[test]
    fn test_parse_frontmatter_not_at_start() {
        let md = r#"Some text first
---
title: "Test"
---

Content
"#;

        let (frontmatter, content) = parse_frontmatter(md);
        assert!(frontmatter.is_none());
        assert_eq!(content, md);
    }

    #[test]
    fn test_parse_yaml_value_quoted_double() {
        assert_eq!(
            parse_yaml_value("\"Hello World\""),
            Some("Hello World".to_string())
        );
        assert_eq!(parse_yaml_value("\"Test\""), Some("Test".to_string()));
    }

    #[test]
    fn test_parse_yaml_value_quoted_single() {
        assert_eq!(
            parse_yaml_value("'Hello World'"),
            Some("Hello World".to_string())
        );
        assert_eq!(parse_yaml_value("'Test'"), Some("Test".to_string()));
    }

    #[test]
    fn test_parse_yaml_value_unquoted() {
        assert_eq!(
            parse_yaml_value("Hello World"),
            Some("Hello World".to_string())
        );
        assert_eq!(parse_yaml_value("Test"), Some("Test".to_string()));
    }

    #[test]
    fn test_parse_yaml_value_empty() {
        assert_eq!(parse_yaml_value(""), None);
        assert_eq!(parse_yaml_value("   "), None);
    }

    #[test]
    fn test_parse_bool_true() {
        assert!(parse_bool("true"));
        assert!(parse_bool("True"));
        assert!(parse_bool("TRUE"));
        assert!(parse_bool("yes"));
        assert!(parse_bool("1"));
        assert!(parse_bool("\"true\""));
        assert!(parse_bool("'true'"));
    }

    #[test]
    fn test_parse_bool_false() {
        assert!(!parse_bool("false"));
        assert!(!parse_bool("False"));
        assert!(!parse_bool("FALSE"));
        assert!(!parse_bool("no"));
        assert!(!parse_bool("0"));
        assert!(!parse_bool("\"false\""));
        assert!(!parse_bool("'false'"));
        assert!(!parse_bool("unknown")); // Default to false
    }

    #[test]
    fn test_parse_frontmatter_with_comments() {
        let md = r#"---
# This is a comment
title: "Test"
# Another comment
skip_toc: true
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Test".to_string()));
        assert_eq!(fm.skip_toc, true);
    }

    #[test]
    fn test_parse_frontmatter_with_empty_lines() {
        let md = r#"---
title: "Test"

skip_toc: true

---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Test".to_string()));
        assert_eq!(fm.skip_toc, true);
    }

    #[test]
    fn test_parse_markdown_with_frontmatter() {
        let md = r#"---
title: "Getting Started"
skip_toc: false
---

# Chapter 1

Content here.
"#;

        let doc = parse_markdown_with_frontmatter(md);
        assert!(doc.frontmatter.is_some());

        let fm = doc.frontmatter.unwrap();
        assert_eq!(fm.title, Some("Getting Started".to_string()));
        assert_eq!(fm.skip_toc, false);

        // Check that markdown was parsed
        assert!(!doc.blocks.is_empty());
        assert!(matches!(
            doc.blocks[0],
            crate::parser::ast::Block::Heading { .. }
        ));
    }

    #[test]
    fn test_parse_markdown_with_frontmatter_no_frontmatter() {
        let md = "# Chapter 1\n\nContent here.";

        let doc = parse_markdown_with_frontmatter(md);
        assert!(doc.frontmatter.is_none());

        // Check that markdown was still parsed
        assert!(!doc.blocks.is_empty());
    }

    #[test]
    fn test_parse_frontmatter_thai_text() {
        let md = r#"---
title: "บทที่ 1"
title_th: "บทที่ 1"
custom: "ค่าทดสอบ"
---

# Chapter 1
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("บทที่ 1".to_string()));
        assert_eq!(fm.title_th, Some("บทที่ 1".to_string()));
        assert_eq!(fm.extra.get("custom"), Some(&"ค่าทดสอบ".to_string()));
    }

    #[test]
    fn test_parse_frontmatter_boolean_variations() {
        let md = r#"---
skip_toc: true
skip_numbering: false
page_break_before: "true"
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert!(fm.skip_toc);
        assert!(!fm.skip_numbering);
        assert!(fm.page_break_before);
    }

    #[test]
    fn test_parse_frontmatter_multiline_value() {
        // Note: Our simple parser doesn't support multiline values
        // This test documents current behavior
        let md = r#"---
title: "Test"
description: This is a long
description that spans
multiple lines
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Test".to_string()));
        // Only first line is captured
        assert_eq!(
            fm.extra.get("description"),
            Some(&"This is a long".to_string())
        );
    }

    #[test]
    fn test_parse_frontmatter_colon_in_value() {
        let md = r#"---
title: "Chapter 1: Introduction"
url: https://example.com:8080
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Chapter 1: Introduction".to_string()));
        assert_eq!(
            fm.extra.get("url"),
            Some(&"https://example.com:8080".to_string())
        );
    }

    #[test]
    fn test_parse_frontmatter_special_characters() {
        let md = r#"---
title: "Test & Demo"
description: "Price: $100, Discount: 20%"
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Test & Demo".to_string()));
        assert_eq!(
            fm.extra.get("description"),
            Some(&"Price: $100, Discount: 20%".to_string())
        );
    }

    #[test]
    fn test_parse_frontmatter_only_unknown_keys() {
        let md = r#"---
custom_field: "value"
another_field: another_value
number_field: 123
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, None);
        assert_eq!(fm.extra.get("custom_field"), Some(&"value".to_string()));
        assert_eq!(
            fm.extra.get("another_field"),
            Some(&"another_value".to_string())
        );
        assert_eq!(fm.extra.get("number_field"), Some(&"123".to_string()));
    }

    #[test]
    fn test_parse_frontmatter_whitespace_handling() {
        let md = r#"---
  title  :  "Test"  
  skip_toc  :  true  
---

Content
"#;

        let (frontmatter, _) = parse_frontmatter(md);
        assert!(frontmatter.is_some());

        let fm = frontmatter.unwrap();
        assert_eq!(fm.title, Some("Test".to_string()));
        assert!(fm.skip_toc);
    }
}
