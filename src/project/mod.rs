//! Project builder for md2docx
//!
//! This module provides a high-level API for building DOCX documents from
//! project directories containing markdown files and configuration.

mod markdown;

use std::path::{Path, PathBuf};

#[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
use crate::config::ProjectConfig;
#[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
use crate::discovery::DiscoveredProject;
use crate::{
    markdown_to_docx_with_templates, DocumentConfig, Error, Language, PlaceholderContext, Result,
    TemplateDir, TemplateSet,
};

pub use markdown::{extract_cover_inside_content, resolve_image_paths, strip_frontmatter};

/// High-level project builder for converting markdown projects to DOCX
///
/// # Example
/// ```rust,ignore
/// use md2docx::project::ProjectBuilder;
///
/// let docx_bytes = ProjectBuilder::from_directory("./docs")?
///     .with_toc(true)
///     .build()?;
/// ```
#[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
pub struct ProjectBuilder {
    base_dir: PathBuf,
    config: ProjectConfig,
    project: DiscoveredProject,
    templates: Option<TemplateSet>,
    toc_override: Option<bool>,
    output_override: Option<PathBuf>,
}

#[cfg(all(feature = "cli", not(target_arch = "wasm32")))]
impl ProjectBuilder {
    /// Create a builder from a directory path
    ///
    /// Automatically detects and loads `md2docx.toml` if present.
    /// Discovers markdown files using the config patterns.
    /// Loads templates from the configured template directory.
    pub fn from_directory(dir: impl AsRef<Path>) -> Result<Self> {
        let base_dir = dir.as_ref().to_path_buf();

        // Load config from md2docx.toml if it exists
        let config_path = base_dir.join("md2docx.toml");
        let config = if config_path.exists() {
            ProjectConfig::from_file(&config_path)?
        } else {
            ProjectConfig::default()
        };

        // Discover project files
        let project = DiscoveredProject::discover_with_config(&base_dir, &config)?;

        // Load templates if configured
        let templates = if let Some(ref template_dir) = config.template.dir {
            let template_path = base_dir.join(template_dir);
            if template_path.exists() {
                let template_dir_obj = TemplateDir::load(&template_path)?;
                Some(template_dir_obj.load_all()?)
            } else {
                None
            }
        } else {
            None
        };

        Ok(Self {
            base_dir,
            config,
            project,
            templates,
            toc_override: None,
            output_override: None,
        })
    }

    /// Override TOC settings from CLI
    pub fn with_toc(mut self, enabled: bool) -> Self {
        self.toc_override = Some(enabled);
        self
    }

    /// Override output path from CLI
    pub fn with_output(mut self, path: PathBuf) -> Self {
        self.output_override = Some(path);
        self
    }

    /// Build the DOCX document and return bytes
    pub fn build(self) -> Result<Vec<u8>> {
        if !self.project.is_valid() {
            return Err(Error::Config(
                "No markdown files found in project directory".into(),
            ));
        }

        // Combine markdown files
        let (combined_markdown, first_content_dir) = self.combine_markdown_files()?;

        // Determine language
        let lang = if self.config.is_thai() {
            Language::Thai
        } else {
            Language::English
        };

        // Build placeholder context
        let placeholder_ctx = self.build_placeholder_context();

        // Build document config
        let doc_config = self.build_document_config(first_content_dir);

        // Change to project directory for relative image paths
        let original_dir = std::env::current_dir()?;
        std::env::set_current_dir(&self.base_dir)?;

        let result = markdown_to_docx_with_templates(
            &combined_markdown,
            lang,
            &doc_config,
            self.templates.as_ref(),
            &placeholder_ctx,
        );

        std::env::set_current_dir(original_dir)?;

        result
    }

    /// Build the DOCX document and write to file
    ///
    /// Returns the path of the output file.
    pub fn build_to_file(self) -> Result<PathBuf> {
        let output_path = self.resolve_output_path();
        let docx_bytes = self.build()?;

        // Create parent directories if needed
        if let Some(parent) = output_path.parent() {
            if !parent.exists() {
                std::fs::create_dir_all(parent)?;
            }
        }

        std::fs::write(&output_path, docx_bytes)?;
        Ok(output_path)
    }

    /// Get the base directory
    pub fn base_dir(&self) -> &Path {
        &self.base_dir
    }

    /// Get the loaded config
    pub fn config(&self) -> &ProjectConfig {
        &self.config
    }

    /// Check if templates are loaded
    pub fn has_templates(&self) -> bool {
        self.templates.is_some()
    }

    /// Get the discovered project
    pub fn project(&self) -> &DiscoveredProject {
        &self.project
    }

    // --- Private helpers ---

    fn resolve_output_path(&self) -> PathBuf {
        if let Some(ref override_path) = self.output_override {
            return override_path.clone();
        }

        if let Some(resolved) = self.config.output.resolve_filename(Some(&self.config)) {
            // Output path is relative to current directory, not input directory
            resolved
        } else {
            self.base_dir.join("output.docx")
        }
    }

    fn combine_markdown_files(&self) -> Result<(String, Option<PathBuf>)> {
        let files = self.project.all_files();
        let mut combined = String::new();
        let mut first_content_dir: Option<PathBuf> = None;

        // Check if using cover template - if so, skip cover.md from main content
        let using_cover_template = self
            .templates
            .as_ref()
            .map(|t| t.has_cover())
            .unwrap_or(false);

        for file_path in files {
            let file_name = file_path
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("unknown");

            // Skip cover.md if using cover template (it's rendered via {{inside}} placeholder)
            if using_cover_template && file_name == "cover.md" {
                continue;
            }

            // Track first content file directory
            if first_content_dir.is_none() {
                if let Some(parent) = file_path.parent() {
                    first_content_dir = Some(parent.to_path_buf());
                }
            }

            let raw_content = std::fs::read_to_string(file_path)?;

            // Strip frontmatter
            let content_without_frontmatter = strip_frontmatter(&raw_content);

            // Resolve image paths
            let content = resolve_image_paths(&content_without_frontmatter, file_path);

            // Add section break between chapters
            if !combined.is_empty() {
                combined.push_str("\n\n---\n\n");
            }

            combined.push_str(&content);
        }

        Ok((combined, first_content_dir))
    }

    fn build_placeholder_context(&self) -> PlaceholderContext {
        let mut ctx = PlaceholderContext::default();
        ctx.set("title", &self.config.document.title);
        ctx.set("subtitle", &self.config.document.subtitle);
        ctx.set("author", &self.config.document.author);
        #[cfg(feature = "cli")]
        ctx.set("date", self.config.date());
        #[cfg(not(feature = "cli"))]
        ctx.set("date", &self.config.document.date);
        ctx.set("version", &self.config.document.version);

        // Pass user-defined extra variables from [document] section
        for (key, value) in self.config.document.extra_as_strings() {
            ctx.set(&key, value);
        }

        // Extract inside content from cover.md if using cover template
        if self
            .templates
            .as_ref()
            .map(|t| t.has_cover())
            .unwrap_or(false)
        {
            if let Some(inside) = extract_cover_inside_content(&self.base_dir) {
                ctx = ctx.with_custom("inside", inside);
            }
        }

        ctx
    }

    fn build_document_config(&self, first_content_dir: Option<PathBuf>) -> DocumentConfig {
        let template_loaded = self.templates.is_some();

        // Load header/footer template if available
        let header_footer_template = if let Some(ref template_dir) = self.config.template.dir {
            let hf_path = self.base_dir.join(template_dir).join("header-footer.docx");
            if hf_path.exists() {
                crate::template::extract::header_footer::extract(&hf_path).ok()
            } else {
                None
            }
        } else {
            None
        };

        // Build font config
        let fonts = Some(crate::docx::ooxml::FontConfig {
            default: if self.config.fonts.default.is_empty() {
                None
            } else {
                Some(self.config.fonts.default.clone())
            },
            code: if self.config.fonts.code.is_empty() {
                None
            } else {
                Some(self.config.fonts.code.clone())
            },
            normal_size: Some(self.config.fonts.normal_based_size * 2),
            normal_color: Some(self.config.fonts.normal_based_color.clone()),
            h1_color: Some(self.config.fonts.h1_based_color.clone()),
            caption_size: Some(self.config.fonts.caption_based_size * 2),
            caption_color: Some(self.config.fonts.caption_based_color.clone()),
            code_size: Some(self.config.fonts.code_based_size * 2),
        });

        // Determine TOC settings
        let mut toc_enabled = self.toc_override.unwrap_or(self.config.toc.enabled);

        // Only disable TOC if pattern is explicitly empty
        if self.config.chapters.pattern.is_empty() {
            toc_enabled = false;
        }

        // Build page config from md2docx.toml settings
        let page_config = {
            use crate::docx::{parse_length_to_twips, PageConfig};

            let width = parse_length_to_twips(&self.config.document.page_width);
            let height = parse_length_to_twips(&self.config.document.page_height);
            let margin_top = parse_length_to_twips(&self.config.document.page_margin_top);
            let margin_bottom = parse_length_to_twips(&self.config.document.page_margin_bottom);
            let margin_left = parse_length_to_twips(&self.config.document.page_margin_left);
            let margin_right = parse_length_to_twips(&self.config.document.page_margin_right);

            // Only create PageConfig if at least one value is set
            if width.is_some()
                || height.is_some()
                || margin_top.is_some()
                || margin_bottom.is_some()
                || margin_left.is_some()
                || margin_right.is_some()
            {
                Some(PageConfig {
                    width,
                    height,
                    margin_top,
                    margin_right,
                    margin_bottom,
                    margin_left,
                    margin_header: None, // Not configured in TOML yet
                    margin_footer: None, // Not configured in TOML yet
                    margin_gutter: None, // Not configured in TOML yet
                })
            } else {
                None
            }
        };

        // Prepare embedded fonts if enabled
        let embedded_fonts = if self.config.fonts.embed {
            let font_dir = if let Some(ref embed_dir) = self.config.fonts.embed_dir {
                self.base_dir.join(embed_dir)
            } else if let Some(ref template_dir) = self.config.template.dir {
                // Default: look for fonts in template/fonts/
                self.base_dir.join(template_dir).join("fonts")
            } else {
                self.base_dir.join("fonts")
            };

            if font_dir.exists() {
                // Collect font names to embed: the default font and code font
                let mut font_names: Vec<&str> = Vec::new();
                if !self.config.fonts.default.is_empty() {
                    font_names.push(&self.config.fonts.default);
                }
                if !self.config.fonts.code.is_empty() {
                    font_names.push(&self.config.fonts.code);
                }

                // If no specific fonts named, embed all fonts found in the directory
                if font_names.is_empty() {
                    match crate::docx::font_embed::scan_font_dir(&font_dir) {
                        Ok(families) => {
                            let all_names: Vec<String> = families.keys().cloned().collect();
                            let name_refs: Vec<&str> = all_names.iter().map(|s| s.as_str()).collect();
                            crate::docx::font_embed::prepare_embedded_fonts(&font_dir, &name_refs)
                                .unwrap_or_default()
                        }
                        Err(_) => Vec::new(),
                    }
                } else {
                    crate::docx::font_embed::prepare_embedded_fonts(&font_dir, &font_names)
                        .unwrap_or_default()
                }
            } else {
                eprintln!(
                    "Warning: Font embed directory not found: {}",
                    font_dir.display()
                );
                Vec::new()
            }
        } else {
            Vec::new()
        };

        DocumentConfig {
            title: self.config.document.title.clone(),
            toc: crate::docx::toc::TocConfig {
                enabled: toc_enabled,
                depth: self.config.toc.depth,
                title: self.config.toc.title.clone(),
                after_cover: self.config.toc.after_cover,
            },
            header_footer_template,
            document_meta: Some(crate::DocumentMeta {
                title: self.config.document.title.clone(),
                subtitle: self.config.document.subtitle.clone(),
                author: self.config.document.author.clone(),
                #[cfg(feature = "cli")]
                date: self.config.date(),
                #[cfg(not(feature = "cli"))]
                date: self.config.document.date.clone(),
            }),
            fonts,
            template_dir: self
                .config
                .template
                .dir
                .as_ref()
                .map(|d| self.base_dir.join(d)),
            process_all_headings: template_loaded,
            base_path: first_content_dir,
            page: page_config,
            embedded_fonts,
            ..DocumentConfig::default()
        }
    }
}
