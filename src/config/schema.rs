//! Configuration schema for md2docx
//!
//! This module defines the structure of `md2docx.toml` configuration files
//! and provides methods to load and parse them.

use serde::Deserialize;
use std::collections::HashMap;
use std::path::{Path, PathBuf};

/// Top-level project configuration from md2docx.toml
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct ProjectConfig {
    pub document: DocumentSection,
    pub template: TemplateSection,
    pub output: OutputSection,
    pub toc: TocSection,
    pub fonts: FontsSection,
    pub code: CodeSection,
    pub chapters: ChaptersSection,
    pub appendices: AppendicesSection,
    pub cover: CoverSection,
}

/// Document metadata section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct DocumentSection {
    pub title: String,
    pub subtitle: String,
    pub author: String,
    pub date: String,     // "auto" or specific date
    pub language: String, // "en" or "th"
    pub version: String,
    pub page_width: String,
    pub page_height: String,
    pub page_margin_top: String,
    pub page_margin_bottom: String,
    pub page_margin_left: String,
    pub page_margin_right: String,
    /// User-defined custom variables (any extra keys in [document])
    /// These are available as {{key}} placeholders in cover templates and output filenames.
    #[serde(flatten)]
    pub extra: HashMap<String, toml::Value>,
}

impl Default for DocumentSection {
    fn default() -> Self {
        Self {
            title: String::new(),
            subtitle: String::new(),
            author: String::new(),
            date: String::new(),
            language: String::new(),
            version: String::new(),
            page_width: "210mm".to_string(),
            page_height: "297mm".to_string(),
            page_margin_top: "25.4mm".to_string(),
            page_margin_bottom: "25.4mm".to_string(),
            page_margin_left: "25.4mm".to_string(),
            page_margin_right: "25.4mm".to_string(),
            extra: HashMap::new(),
        }
    }
}

impl DocumentSection {
    /// Get all extra (user-defined) variables as string key-value pairs.
    /// Non-string TOML values are converted to their display representation.
    pub fn extra_as_strings(&self) -> HashMap<String, String> {
        self.extra
            .iter()
            .map(|(k, v)| {
                let s = match v {
                    toml::Value::String(s) => s.clone(),
                    toml::Value::Integer(i) => i.to_string(),
                    toml::Value::Float(f) => f.to_string(),
                    toml::Value::Boolean(b) => b.to_string(),
                    toml::Value::Datetime(d) => d.to_string(),
                    other => other.to_string(),
                };
                (k.clone(), s)
            })
            .collect()
    }
}

/// Template configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct TemplateSection {
    /// Template directory containing cover.docx, table.docx, etc.
    pub dir: Option<PathBuf>,
    /// Validate template has required styles
    pub validate: bool,
}

/// Output file configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct OutputSection {
    pub file: Option<PathBuf>,
}

impl OutputSection {
    /// Resolve filename by expanding placeholders like {{currenttime:FORMAT}}, {{title}}, {{author}}, etc.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn resolve_filename(&self, project_config: Option<&ProjectConfig>) -> Option<PathBuf> {
        self.file.as_ref().map(|p| {
            let path_str = p.to_string_lossy();
            let mut result = path_str.to_string();

            // Expand {{currenttime:FORMAT}} placeholders
            if result.contains("{{currenttime:") {
                result = expand_currenttime_placeholder(&result);
            }

            // Expand document variable placeholders if project_config is provided
            if let Some(config) = project_config {
                result = expand_document_placeholders(&result, &config.document);
            }

            PathBuf::from(result)
        })
    }
}

/// Expand {{currenttime:FORMAT}} to actual timestamp using LOCAL device time
/// Supports native chrono format strings (e.g., "%Y%m%d_%H%M%S")
#[cfg(not(target_arch = "wasm32"))]
fn expand_currenttime_placeholder(template: &str) -> String {
    use chrono::Local;

    let now = Local::now();

    let mut result = template.to_string();
    while let Some(start) = result.find("{{currenttime:") {
        if let Some(end) = result[start..].find("}}") {
            let full_placeholder = &result[start..start + end + 2];
            let format_str = &result[start + 14..start + end];

            // Use chrono's native strftime formatting
            let formatted = now.format(format_str).to_string();

            result = result.replace(full_placeholder, &formatted);
        } else {
            break;
        }
    }
    result
}

/// Expand document variable placeholders like {{title}}, {{author}}, {{version}}, etc.
#[cfg(not(target_arch = "wasm32"))]
fn expand_document_placeholders(template: &str, document: &DocumentSection) -> String {
    let mut result = template.to_string();

    // Define all available placeholders and their values
    let placeholders = [
        ("{{title}}", sanitize_filename(&document.title)),
        ("{{author}}", sanitize_filename(&document.author)),
        ("{{version}}", sanitize_filename(&document.version)),
        ("{{subtitle}}", sanitize_filename(&document.subtitle)),
        ("{{language}}", document.language.clone()),
        ("{{date}}", sanitize_filename(&document.date)),
    ];

    for (placeholder, value) in &placeholders {
        result = result.replace(placeholder, value);
    }

    // Also expand any user-defined extra variables from [document]
    for (key, value) in document.extra_as_strings() {
        let placeholder = format!("{{{{{}}}}}", key);
        result = result.replace(&placeholder, &sanitize_filename(&value));
    }

    result
}

/// Sanitize a string to be safe for use in filenames
/// Removes or replaces characters that are invalid in filenames
#[cfg(not(target_arch = "wasm32"))]
fn sanitize_filename(input: &str) -> String {
    if input.is_empty() {
        return "unknown".to_string();
    }

    // Characters that are invalid in Windows filenames
    let invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*'];

    let mut result: String = input
        .chars()
        .map(|c| if invalid_chars.contains(&c) { '_' } else { c })
        .collect();

    // Trim whitespace and limit length
    result = result.trim().to_string();
    if result.len() > 100 {
        // Find the nearest char boundary at or before byte 100
        let mut end = 100;
        while !result.is_char_boundary(end) {
            end -= 1;
        }
        result = result[..end].to_string();
    }

    // If empty after sanitization, return "unknown"
    if result.is_empty() {
        result = "unknown".to_string();
    }

    result
}

/// Table of contents configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct TocSection {
    pub enabled: bool,
    pub depth: u8,
    pub title: String,
    pub after_cover: bool, // If true, TOC comes after cover content
}

impl Default for TocSection {
    fn default() -> Self {
        Self {
            enabled: false,
            depth: 3,
            title: "Table of Contents".to_string(),
            after_cover: true,
        }
    }
}

/// Font configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct FontsSection {
    pub default: String,
    pub code: String,
    pub normal_based_size: u32,
    pub normal_based_color: String,
    pub h1_based_color: String,
    pub caption_based_size: u32,
    pub caption_based_color: String,
    pub code_based_size: u32,
    /// Enable font embedding in the generated DOCX
    pub embed: bool,
    /// Directory containing .ttf/.otf font files to embed
    pub embed_dir: Option<PathBuf>,
}

impl Default for FontsSection {
    fn default() -> Self {
        Self {
            default: String::new(),
            code: String::new(),
            normal_based_size: 11,
            normal_based_color: "#000000".to_string(),
            h1_based_color: "#2F5496".to_string(),
            caption_based_size: 9,
            caption_based_color: "#000000".to_string(),
            code_based_size: 10,
            embed: false,
            embed_dir: None,
        }
    }
}

/// Code block configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct CodeSection {
    pub theme: String,
    pub show_filename: bool,
    pub show_line_numbers: bool,
    pub source_root: Option<PathBuf>,
}

impl Default for CodeSection {
    fn default() -> Self {
        Self {
            theme: "light".to_string(),
            show_filename: true,
            show_line_numbers: false,
            source_root: None,
        }
    }
}

/// Chapters configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct ChaptersSection {
    pub pattern: String,
    pub sort: String,
}

impl Default for ChaptersSection {
    fn default() -> Self {
        Self {
            pattern: "ch*_*.md".to_string(),
            sort: "numeric".to_string(),
        }
    }
}

/// Appendices configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct AppendicesSection {
    pub pattern: String,
    pub prefix: String,
}

impl Default for AppendicesSection {
    fn default() -> Self {
        Self {
            pattern: "ap*_*.md".to_string(),
            prefix: "Appendix".to_string(),
        }
    }
}

/// Cover page configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct CoverSection {
    pub file: Option<PathBuf>,
    pub title: Option<String>,
    pub subtitle: Option<String>,
    pub date: Option<String>,
}

impl ProjectConfig {
    /// Load config from a TOML file
    #[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
    pub fn from_file(path: &Path) -> crate::Result<Self> {
        let content = std::fs::read_to_string(path)?;
        Self::parse_toml(&content)
    }

    /// Parse config from a TOML string
    #[cfg(feature = "cli")]
    pub fn parse_toml(toml_content: &str) -> crate::Result<Self> {
        toml::from_str(toml_content)
            .map_err(|e| crate::Error::Config(format!("Failed to parse config: {}", e)))
    }

    /// Get the effective language (default to "en" if not specified)
    pub fn language(&self) -> &str {
        let lang = self.document.language.trim();
        if lang.is_empty() {
            "en"
        } else {
            lang
        }
    }

    /// Check if Thai language is configured
    pub fn is_thai(&self) -> bool {
        matches!(self.language().to_lowercase().as_str(), "th" | "thai")
    }

    /// Get the effective date string
    #[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
    pub fn date(&self) -> String {
        if self.document.date == "auto" {
            // Use expand_currenttime_placeholder to get YYYY-MM-DD
            expand_currenttime_placeholder("{{currenttime:YYYY-MM-DD}}")
        } else {
            self.document.date.clone()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[cfg(feature = "cli")]
    fn test_parse_minimal_config() {
        let toml = r##"
[document]
title = "Test Document"
"##;

        let config = ProjectConfig::parse_toml(toml).unwrap();
        assert_eq!(config.document.title, "Test Document");
        assert_eq!(config.document.language, ""); // Default empty
        assert_eq!(config.toc.enabled, false); // Default
    }

    #[test]
    #[cfg(feature = "cli")]
    fn test_parse_full_config() {
        let toml = r##"
[document]
title = "คู่มือการใช้งาน"
subtitle = "แปลง Markdown เป็น DOCX"
author = "ทีมพัฒนา"
date = "auto"
language = "th"
version = "1.0.0"
page_width = "210mm"
page_height = "297mm"

[template]
dir = "template"
validate = true

[output]
file = "output/manual.docx"

[toc]
enabled = true
depth = 3
title = "สารบัญ"

[fonts]
default = "TH Sarabun New"
code = "Consolas"
normal_based_size = 14
h1_based_color = "#000080"

[code]
theme = "light"
show_filename = true
show_line_numbers = false
source_root = "../src"

[chapters]
pattern = "ch*_*.md"
sort = "numeric"

[appendices]
pattern = "ap*_*.md"
prefix = "Appendix"

[cover]
file = "cover.md"
title = "Custom Title"
subtitle = "Custom Subtitle"
date = "2025-01-28"
"##;

        let config = ProjectConfig::parse_toml(toml).unwrap();

        // Document section
        assert_eq!(config.document.title, "คู่มือการใช้งาน");
        assert_eq!(config.document.subtitle, "แปลง Markdown เป็น DOCX");
        assert_eq!(config.document.author, "ทีมพัฒนา");
        assert_eq!(config.document.date, "auto");
        assert_eq!(config.document.language, "th");
        assert_eq!(config.document.version, "1.0.0");
        assert_eq!(config.document.page_width, "210mm");

        // Template section
        assert_eq!(config.template.dir, Some(PathBuf::from("template")));
        assert_eq!(config.template.validate, true);

        // Output section
        assert_eq!(
            config.output.file,
            Some(PathBuf::from("output/manual.docx"))
        );

        // TOC section
        assert_eq!(config.toc.enabled, true);
        assert_eq!(config.toc.depth, 3);
        assert_eq!(config.toc.title, "สารบัญ");

        // Fonts section
        assert_eq!(config.fonts.default, "TH Sarabun New");
        assert_eq!(config.fonts.code, "Consolas");
        assert_eq!(config.fonts.normal_based_size, 14);
        assert_eq!(config.fonts.h1_based_color, "#000080");

        // Code section
        assert_eq!(config.code.theme, "light");
        assert_eq!(config.code.show_filename, true);
        assert_eq!(config.code.show_line_numbers, false);
        assert_eq!(config.code.source_root, Some(PathBuf::from("../src")));

        // Chapters section
        assert_eq!(config.chapters.pattern, "ch*_*.md");
        assert_eq!(config.chapters.sort, "numeric");

        // Appendices section
        assert_eq!(config.appendices.pattern, "ap*_*.md");
        assert_eq!(config.appendices.prefix, "Appendix");

        // Cover section
        assert_eq!(config.cover.file, Some(PathBuf::from("cover.md")));
        assert_eq!(config.cover.title, Some("Custom Title".to_string()));
        assert_eq!(config.cover.subtitle, Some("Custom Subtitle".to_string()));
        assert_eq!(config.cover.date, Some("2025-01-28".to_string()));
    }

    #[test]
    #[cfg(feature = "cli")]
    fn test_parse_thai_example_config() {
        let toml = r##"
[document]
title = "คู่มือการใช้งาน md2docx"
subtitle = "แปลง Markdown เป็น DOCX อย่างมืออาชีพ"
author = "ทีมพัฒนา md2docx"
date = "auto"
language = "th"

[template]
dir = "template"

[output]
file = "output/คู่มือ-md2docx.docx"

[toc]
enabled = true
depth = 3

[fonts]
default = "TH Sarabun New"
code = "Consolas"

[code]
theme = "light"
show_filename = true
show_line_numbers = false

[chapters]
pattern = "ch*_*.md"
sort = "numeric"

[appendices]
pattern = "ap*_*.md"
"##;

        let config = ProjectConfig::parse_toml(toml).unwrap();

        assert_eq!(config.document.title, "คู่มือการใช้งาน md2docx");
        assert_eq!(config.document.language, "th");
        assert_eq!(config.toc.enabled, true);
        assert_eq!(config.toc.depth, 3);
        assert_eq!(config.fonts.default, "TH Sarabun New");
        assert_eq!(config.chapters.pattern, "ch*_*.md");
    }

    #[test]
    #[cfg(feature = "cli")]
    fn test_default_values() {
        let toml = "";
        let config = ProjectConfig::parse_toml(toml).unwrap();

        // Check defaults
        assert_eq!(config.document.title, "");
        assert_eq!(config.document.subtitle, "");
        assert_eq!(config.document.author, "");
        assert_eq!(config.document.date, "");
        assert_eq!(config.document.language, "");
        assert_eq!(config.document.page_width, "210mm");

        assert_eq!(config.template.dir, None);
        assert_eq!(config.template.validate, false);

        assert_eq!(config.output.file, None);

        assert_eq!(config.toc.enabled, false);
        assert_eq!(config.toc.depth, 3);
        assert_eq!(config.toc.title, "Table of Contents".to_string());
        assert_eq!(config.toc.after_cover, true);

        assert_eq!(config.fonts.default, "");
        assert_eq!(config.fonts.code, "");
        assert_eq!(config.fonts.normal_based_size, 11);

        assert_eq!(config.code.theme, "light");
        assert_eq!(config.code.show_filename, true);
        assert_eq!(config.code.show_line_numbers, false);
        assert_eq!(config.code.source_root, None);

        assert_eq!(config.chapters.pattern, "ch*_*.md");
        assert_eq!(config.chapters.sort, "numeric");

        assert_eq!(config.appendices.pattern, "ap*_*.md");
        assert_eq!(config.appendices.prefix, "Appendix");

        assert_eq!(config.cover.file, None);
        assert_eq!(config.cover.title, None);
        assert_eq!(config.cover.subtitle, None);
        assert_eq!(config.cover.date, None);
    }

    #[test]
    #[cfg(feature = "cli")]
    fn test_invalid_toml() {
        let toml = r#"
[document
title = "Missing closing bracket"
"#;

        let result = ProjectConfig::parse_toml(toml);
        assert!(result.is_err());
        if let Err(crate::Error::Config(msg)) = result {
            assert!(msg.contains("Failed to parse config"));
        } else {
            panic!("Expected Config error");
        }
    }

    #[test]
    fn test_language_helper() {
        let mut config = ProjectConfig::default();

        // Empty language defaults to "en"
        assert_eq!(config.language(), "en");
        assert!(!config.is_thai());

        // Thai language
        config.document.language = "th".to_string();
        assert_eq!(config.language(), "th");
        assert!(config.is_thai());

        // Thai language (uppercase)
        config.document.language = "TH".to_string();
        assert_eq!(config.language(), "TH");
        assert!(config.is_thai());

        // Thai language (full name)
        config.document.language = "thai".to_string();
        assert_eq!(config.language(), "thai");
        assert!(config.is_thai());

        // English language
        config.document.language = "en".to_string();
        assert_eq!(config.language(), "en");
        assert!(!config.is_thai());
    }

    #[test]
    #[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
    fn test_date_helper() {
        let mut config = ProjectConfig::default();

        // Auto date
        config.document.date = "auto".to_string();
        let date = config.date();
        // Should be in YYYY-MM-DD format
        assert!(date.len() == 10);
        assert!(date.contains('-'));

        // Specific date
        config.document.date = "2025-01-28".to_string();
        assert_eq!(config.date(), "2025-01-28");

        // Empty date
        config.document.date = "".to_string();
        assert_eq!(config.date(), "");
    }

    #[test]
    #[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
    fn test_resolve_filename() {
        let mut output = OutputSection::default();
        output.file = Some(PathBuf::from("output-{{currenttime:YYYYMMDD}}.docx"));

        let resolved = output.resolve_filename(None).unwrap();
        let resolved_str = resolved.to_string_lossy();

        assert!(resolved_str.starts_with("output-"));
        assert!(resolved_str.ends_with(".docx"));
        assert_eq!(resolved_str.len(), "output-20250101.docx".len());
    }

    #[test]
    #[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
    fn test_resolve_filename_with_document_vars() {
        use super::ProjectConfig;

        let mut project_config = ProjectConfig::default();
        project_config.document.title = "My Document".to_string();
        project_config.document.author = "John Doe".to_string();
        project_config.document.version = "1.2.3".to_string();

        let mut output = OutputSection::default();
        output.file = Some(PathBuf::from("{{title}}-v{{version}}-{{author}}.docx"));

        let resolved = output.resolve_filename(Some(&project_config)).unwrap();
        let resolved_str = resolved.to_string_lossy();

        assert_eq!(resolved_str, "My Document-v1.2.3-John Doe.docx");
    }

    #[test]
    #[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
    fn test_sanitize_filename() {
        assert_eq!(sanitize_filename("Hello World"), "Hello World");
        assert_eq!(sanitize_filename("Test: File"), "Test_ File");
        assert_eq!(sanitize_filename("File/Path"), "File_Path");
        assert_eq!(sanitize_filename(""), "unknown");
        assert_eq!(sanitize_filename("   "), "unknown");
    }
}
