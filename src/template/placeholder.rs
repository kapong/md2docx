//! Placeholder replacement system for templates
//!
//! This module provides functionality to replace placeholders in template
//! content with actual values from frontmatter or configuration.
//!
//! # Supported Placeholders
//!
//! - `{{title}}` - Document title
//! - `{{subtitle}}` - Document subtitle
//! - `{{author}}` - Document author
//! - `{{date}}` - Document date
//! - `{{version}}` - Document version
//! - `{{chapter}}` - Current chapter name
//! - `{{page}}` - Current page number
//! - `{{total}}` - Total pages
//! - `{{custom_key}}` - Any custom field from frontmatter
//!
//! # Example
//!
//! ```rust
//! use md2docx::template::{PlaceholderContext, replace_placeholders};
//!
//! let ctx = PlaceholderContext {
//!     title: "My Document".to_string(),
//!     author: "John Doe".to_string(),
//!     ..Default::default()
//! };
//!
//! let result = replace_placeholders("{{title}} by {{author}}", &ctx);
//! assert_eq!(result, "My Document by John Doe");
//! ```

use std::collections::HashMap;

/// Context for placeholder replacement
///
/// Contains all available values that can be used to replace placeholders
/// in template content.
#[derive(Debug, Clone, Default)]
pub struct PlaceholderContext {
    /// Document title
    pub title: String,
    /// Document subtitle
    pub subtitle: String,
    /// Document author
    pub author: String,
    /// Document date
    pub date: String,
    /// Document version
    pub version: String,
    /// Current chapter name
    pub chapter: String,
    /// Current page number
    pub page: String,
    /// Total pages
    pub total: String,
    /// Custom fields from frontmatter
    pub custom: HashMap<String, String>,
}

impl PlaceholderContext {
    /// Create a new placeholder context with basic fields
    pub fn new(title: impl Into<String>, author: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            author: author.into(),
            ..Default::default()
        }
    }

    /// Add a custom field
    pub fn with_custom(mut self, key: impl Into<String>, value: impl Into<String>) -> Self {
        self.custom.insert(key.into(), value.into());
        self
    }

    /// Set the subtitle
    pub fn with_subtitle(mut self, subtitle: impl Into<String>) -> Self {
        self.subtitle = subtitle.into();
        self
    }

    /// Set the date
    pub fn with_date(mut self, date: impl Into<String>) -> Self {
        self.date = date.into();
        self
    }

    /// Set the version
    pub fn with_version(mut self, version: impl Into<String>) -> Self {
        self.version = version.into();
        self
    }

    /// Set the chapter
    pub fn with_chapter(mut self, chapter: impl Into<String>) -> Self {
        self.chapter = chapter.into();
        self
    }

    /// Set the page number
    pub fn with_page(mut self, page: impl Into<String>) -> Self {
        self.page = page.into();
        self
    }

    /// Set the total pages
    pub fn with_total(mut self, total: impl Into<String>) -> Self {
        self.total = total.into();
        self
    }

    /// Get a value by key (checks standard fields first, then custom)
    pub fn get(&self, key: &str) -> Option<&str> {
        match key {
            "title" => Some(&self.title),
            "subtitle" => Some(&self.subtitle),
            "author" => Some(&self.author),
            "date" => Some(&self.date),
            "version" => Some(&self.version),
            "chapter" => Some(&self.chapter),
            "page" => Some(&self.page),
            "total" => Some(&self.total),
            _ => self.custom.get(key).map(|s| s.as_str()),
        }
    }
}

/// Replace placeholders in content with values from context
///
/// Placeholders are in the format `{{key}}`. Unknown placeholders
/// are left as-is.
///
/// # Arguments
/// * `content` - The content containing placeholders
/// * `ctx` - The context containing replacement values
///
/// # Returns
/// The content with placeholders replaced
///
/// # Example
/// ```rust
/// use md2docx::template::{PlaceholderContext, replace_placeholders};
///
/// let ctx = PlaceholderContext {
///     title: "Hello".to_string(),
///     author: "World".to_string(),
///     ..Default::default()
/// };
///
/// let result = replace_placeholders("{{title}} {{author}}!", &ctx);
/// assert_eq!(result, "Hello World!");
/// ```
pub fn replace_placeholders(content: &str, ctx: &PlaceholderContext) -> String {
    let mut result = content.to_string();

    // Find all placeholders {{key}}
    let placeholder_regex = regex::Regex::new(r"\{\{(\w+)\}\}").unwrap();

    // Replace each placeholder
    for cap in placeholder_regex.captures_iter(content) {
        let full_match = cap.get(0).unwrap().as_str();
        let key = cap.get(1).unwrap().as_str();

        if let Some(value) = ctx.get(key) {
            result = result.replace(full_match, value);
        }
        // If key not found, leave placeholder as-is
    }

    result
}

/// Check if content contains any placeholders
///
/// # Arguments
/// * `content` - The content to check
///
/// # Returns
/// `true` if the content contains at least one placeholder
///
/// # Example
/// ```rust
/// use md2docx::template::has_placeholders;
///
/// assert!(has_placeholders("{{title}}"));
/// assert!(!has_placeholders("No placeholders here"));
/// ```
pub fn has_placeholders(content: &str) -> bool {
    content.contains("{{") && content.contains("}}")
}

/// Extract all unique placeholder keys from content
///
/// # Arguments
/// * `content` - The content to extract placeholders from
///
/// # Returns
/// A vector of unique placeholder keys (without the braces)
///
/// # Example
/// ```rust
/// use md2docx::template::extract_placeholders;
///
/// let keys = extract_placeholders("{{title}} by {{author}} on {{date}}");
/// assert_eq!(keys, vec!["title", "author", "date"]);
/// ```
pub fn extract_placeholders(content: &str) -> Vec<String> {
    let mut keys = Vec::new();
    let placeholder_regex = regex::Regex::new(r"\{\{(\w+)\}\}").unwrap();

    for cap in placeholder_regex.captures_iter(content) {
        let key = cap.get(1).unwrap().as_str().to_string();
        if !keys.contains(&key) {
            keys.push(key);
        }
    }

    keys
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_replace_placeholders() {
        let ctx = PlaceholderContext {
            title: "My Document".to_string(),
            author: "John Doe".to_string(),
            date: "2025-01-28".to_string(),
            ..Default::default()
        };

        let result = replace_placeholders("{{title}} by {{author}} on {{date}}", &ctx);
        assert_eq!(result, "My Document by John Doe on 2025-01-28");
    }

    #[test]
    fn test_replace_unknown_placeholder() {
        let ctx = PlaceholderContext {
            title: "Hello".to_string(),
            ..Default::default()
        };

        let result = replace_placeholders("{{title}} {{unknown}}", &ctx);
        assert_eq!(result, "Hello {{unknown}}");
    }

    #[test]
    fn test_replace_no_placeholders() {
        let ctx = PlaceholderContext::default();

        let result = replace_placeholders("No placeholders here", &ctx);
        assert_eq!(result, "No placeholders here");
    }

    #[test]
    fn test_replace_custom_fields() {
        let ctx = PlaceholderContext::default()
            .with_custom("department", "Engineering")
            .with_custom("project", "Alpha");

        let result = replace_placeholders("{{department}} - {{project}}", &ctx);
        assert_eq!(result, "Engineering - Alpha");
    }

    #[test]
    fn test_has_placeholders() {
        assert!(has_placeholders("{{title}}"));
        assert!(has_placeholders("Hello {{name}}!"));
        assert!(!has_placeholders("No placeholders"));
        assert!(!has_placeholders(""));
    }

    #[test]
    fn test_extract_placeholders() {
        let keys = extract_placeholders("{{title}} by {{author}} on {{date}}");
        assert_eq!(keys, vec!["title", "author", "date"]);
    }

    #[test]
    fn test_extract_duplicate_placeholders() {
        let keys = extract_placeholders("{{title}} and {{title}} again");
        assert_eq!(keys, vec!["title"]);
    }

    #[test]
    fn test_placeholder_context_builder() {
        let ctx = PlaceholderContext::new("Title", "Author")
            .with_subtitle("Subtitle")
            .with_date("2025-01-28")
            .with_version("1.0")
            .with_custom("key", "value");

        assert_eq!(ctx.title, "Title");
        assert_eq!(ctx.author, "Author");
        assert_eq!(ctx.subtitle, "Subtitle");
        assert_eq!(ctx.date, "2025-01-28");
        assert_eq!(ctx.version, "1.0");
        assert_eq!(ctx.get("key"), Some("value"));
    }

    #[test]
    fn test_placeholder_context_get() {
        let ctx = PlaceholderContext {
            title: "Test".to_string(),
            custom: {
                let mut map = HashMap::new();
                map.insert("custom_key".to_string(), "custom_value".to_string());
                map
            },
            ..Default::default()
        };

        assert_eq!(ctx.get("title"), Some("Test"));
        assert_eq!(ctx.get("custom_key"), Some("custom_value"));
        assert_eq!(ctx.get("unknown"), None);
    }
}
