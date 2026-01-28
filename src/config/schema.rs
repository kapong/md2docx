//! Configuration schema for md2docx
//!
//! This module defines the structure of `md2docx.toml` configuration files
//! and provides methods to load and parse them.

use serde::Deserialize;
use std::path::{Path, PathBuf};

/// Top-level project configuration from md2docx.toml
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct ProjectConfig {
    pub document: DocumentSection,
    pub template: TemplateSection,
    pub output: OutputSection,
    pub toc: TocSection,
    pub page_numbers: PageNumbersSection,
    pub fonts: FontsSection,
    pub code: CodeSection,
    pub images: ImagesSection,
    pub chapters: ChaptersSection,
    pub appendices: AppendicesSection,
    pub header: HeaderSection,
    pub footer: FooterSection,
    pub cover: CoverSection,
}

/// Document metadata section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct DocumentSection {
    pub title: String,
    pub subtitle: String,
    pub author: String,
    pub date: String,     // "auto" or specific date
    pub language: String, // "en" or "th"
}

/// Template configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct TemplateSection {
    pub file: Option<PathBuf>,
    pub validate: bool,
}

/// Output file configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct OutputSection {
    pub file: Option<PathBuf>,
}

/// Table of contents configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct TocSection {
    pub enabled: bool,
    pub depth: u8,
    pub title: String,
}

impl Default for TocSection {
    fn default() -> Self {
        Self {
            enabled: false,
            depth: 3,
            title: "Table of Contents".to_string(),
        }
    }
}

/// Page numbering configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct PageNumbersSection {
    pub enabled: bool,
    pub skip_cover: bool,
    pub skip_chapter_first: bool,
    pub format: String,
}

impl Default for PageNumbersSection {
    fn default() -> Self {
        Self {
            enabled: true,
            skip_cover: true,
            skip_chapter_first: false,
            format: "{n}".to_string(),
        }
    }
}

/// Font configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct FontsSection {
    pub default: String,
    pub thai: String,
    pub code: String,
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

/// Image configuration section
#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct ImagesSection {
    pub max_width: String,
    pub auto_caption: bool,
    pub default_dpi: u32,
}

impl Default for ImagesSection {
    fn default() -> Self {
        Self {
            max_width: "100%".to_string(),
            auto_caption: true,
            default_dpi: 150,
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

/// Header configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct HeaderSection {
    pub left: String,
    pub center: String,
    pub right: String,
    pub skip_cover: bool,
}

/// Footer configuration section
#[derive(Debug, Clone, Default, Deserialize)]
#[serde(default)]
pub struct FooterSection {
    pub left: String,
    pub center: String,
    pub right: String,
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
            // Return current date in ISO format
            use std::time::{SystemTime, UNIX_EPOCH};
            let duration = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default();
            let secs = duration.as_secs();
            // Simple date calculation from Unix timestamp
            let days_since_epoch = secs / 86400;
            // Unix epoch 1970-01-01 was a Thursday (day 4)
            let year = 1970 + (days_since_epoch / 365);
            let remaining_days = days_since_epoch % 365;
            // Approximate month/day (simplified)
            format!(
                "{}-{:02}-{:02}",
                year,
                1 + (remaining_days / 30) % 12,
                1 + remaining_days % 30
            )
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
        let toml = r#"
[document]
title = "Test Document"
"#;

        let config = ProjectConfig::parse_toml(toml).unwrap();
        assert_eq!(config.document.title, "Test Document");
        assert_eq!(config.document.language, ""); // Default empty
        assert_eq!(config.toc.enabled, false); // Default
        assert_eq!(config.page_numbers.enabled, true); // Default
    }

    #[test]
    #[cfg(feature = "cli")]
    fn test_parse_full_config() {
        let toml = r#"
[document]
title = "คู่มือการใช้งาน"
subtitle = "แปลง Markdown เป็น DOCX"
author = "ทีมพัฒนา"
date = "auto"
language = "th"

[template]
file = "custom-reference.docx"
validate = true

[output]
file = "output/manual.docx"

[toc]
enabled = true
depth = 3

[page_numbers]
enabled = true
skip_cover = true
skip_chapter_first = false
format = "{n} of {total}"

[fonts]
default = "TH Sarabun New"
thai = "TH Sarabun New"
code = "Consolas"

[code]
theme = "light"
show_filename = true
show_line_numbers = false
source_root = "../src"

[images]
max_width = "100%"
auto_caption = true
default_dpi = 150

[chapters]
pattern = "ch*_*.md"
sort = "numeric"

[appendices]
pattern = "ap*_*.md"
prefix = "Appendix"

[header]
left = "{title}"
center = ""
right = "{chapter}"
skip_cover = true

[footer]
left = ""
center = "{page}"
right = ""

[cover]
file = "cover.md"
title = "Custom Title"
subtitle = "Custom Subtitle"
date = "2025-01-28"
"#;

        let config = ProjectConfig::parse_toml(toml).unwrap();

        // Document section
        assert_eq!(config.document.title, "คู่มือการใช้งาน");
        assert_eq!(config.document.subtitle, "แปลง Markdown เป็น DOCX");
        assert_eq!(config.document.author, "ทีมพัฒนา");
        assert_eq!(config.document.date, "auto");
        assert_eq!(config.document.language, "th");

        // Template section
        assert_eq!(
            config.template.file,
            Some(PathBuf::from("custom-reference.docx"))
        );
        assert_eq!(config.template.validate, true);

        // Output section
        assert_eq!(
            config.output.file,
            Some(PathBuf::from("output/manual.docx"))
        );

        // TOC section
        assert_eq!(config.toc.enabled, true);
        assert_eq!(config.toc.depth, 3);

        // Page numbers section
        assert_eq!(config.page_numbers.enabled, true);
        assert_eq!(config.page_numbers.skip_cover, true);
        assert_eq!(config.page_numbers.skip_chapter_first, false);
        assert_eq!(config.page_numbers.format, "{n} of {total}");

        // Fonts section
        assert_eq!(config.fonts.default, "TH Sarabun New");
        assert_eq!(config.fonts.thai, "TH Sarabun New");
        assert_eq!(config.fonts.code, "Consolas");

        // Code section
        assert_eq!(config.code.theme, "light");
        assert_eq!(config.code.show_filename, true);
        assert_eq!(config.code.show_line_numbers, false);
        assert_eq!(config.code.source_root, Some(PathBuf::from("../src")));

        // Images section
        assert_eq!(config.images.max_width, "100%");
        assert_eq!(config.images.auto_caption, true);
        assert_eq!(config.images.default_dpi, 150);

        // Chapters section
        assert_eq!(config.chapters.pattern, "ch*_*.md");
        assert_eq!(config.chapters.sort, "numeric");

        // Appendices section
        assert_eq!(config.appendices.pattern, "ap*_*.md");
        assert_eq!(config.appendices.prefix, "Appendix");

        // Header section
        assert_eq!(config.header.left, "{title}");
        assert_eq!(config.header.center, "");
        assert_eq!(config.header.right, "{chapter}");
        assert_eq!(config.header.skip_cover, true);

        // Footer section
        assert_eq!(config.footer.left, "");
        assert_eq!(config.footer.center, "{page}");
        assert_eq!(config.footer.right, "");

        // Cover section
        assert_eq!(config.cover.file, Some(PathBuf::from("cover.md")));
        assert_eq!(config.cover.title, Some("Custom Title".to_string()));
        assert_eq!(config.cover.subtitle, Some("Custom Subtitle".to_string()));
        assert_eq!(config.cover.date, Some("2025-01-28".to_string()));
    }

    #[test]
    #[cfg(feature = "cli")]
    fn test_parse_thai_example_config() {
        let toml = r#"
[document]
title = "คู่มือการใช้งาน md2docx"
subtitle = "แปลง Markdown เป็น DOCX อย่างมืออาชีพ"
author = "ทีมพัฒนา md2docx"
date = "auto"
language = "th"

[template]
file = "custom-reference.docx"

[output]
file = "output/คู่มือ-md2docx.docx"

[toc]
enabled = true
depth = 3

[page_numbers]
enabled = true
skip_cover = true

[fonts]
default = "TH Sarabun New"
code = "Consolas"

[code]
theme = "light"
show_filename = true
show_line_numbers = false

[images]
max_width = "100%"
auto_caption = true

[chapters]
pattern = "ch*_*.md"
sort = "numeric"

[appendices]
pattern = "ap*_*.md"
"#;

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

        assert_eq!(config.template.file, None);
        assert_eq!(config.template.validate, false);

        assert_eq!(config.output.file, None);

        assert_eq!(config.toc.enabled, false);
        assert_eq!(config.toc.depth, 3);

        assert_eq!(config.page_numbers.enabled, true);
        assert_eq!(config.page_numbers.skip_cover, true);
        assert_eq!(config.page_numbers.skip_chapter_first, false);
        assert_eq!(config.page_numbers.format, "{n}");

        assert_eq!(config.fonts.default, "");
        assert_eq!(config.fonts.thai, "");
        assert_eq!(config.fonts.code, "");

        assert_eq!(config.code.theme, "light");
        assert_eq!(config.code.show_filename, true);
        assert_eq!(config.code.show_line_numbers, false);
        assert_eq!(config.code.source_root, None);

        assert_eq!(config.images.max_width, "100%");
        assert_eq!(config.images.auto_caption, true);
        assert_eq!(config.images.default_dpi, 150);

        assert_eq!(config.chapters.pattern, "ch*_*.md");
        assert_eq!(config.chapters.sort, "numeric");

        assert_eq!(config.appendices.pattern, "ap*_*.md");
        assert_eq!(config.appendices.prefix, "Appendix");

        assert_eq!(config.header.left, "");
        assert_eq!(config.header.center, "");
        assert_eq!(config.header.right, "");
        assert_eq!(config.header.skip_cover, false);

        assert_eq!(config.footer.left, "");
        assert_eq!(config.footer.center, "");
        assert_eq!(config.footer.right, "");

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
    #[cfg(feature = "cli")]
    fn test_partial_config() {
        let toml = r#"
[document]
title = "Partial Config"

[toc]
enabled = true
"#;

        let config = ProjectConfig::parse_toml(toml).unwrap();

        assert_eq!(config.document.title, "Partial Config");
        assert_eq!(config.toc.enabled, true);
        // Other sections should have defaults
        assert_eq!(config.page_numbers.enabled, true);
        assert_eq!(config.code.theme, "light");
    }
}
