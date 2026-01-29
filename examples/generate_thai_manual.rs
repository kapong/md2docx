//! Example: Generate Thai Manual from Markdown files with Template Directory
//!
//! This example demonstrates how to:
//! - Read configuration from md2docx.toml
//! - Load templates from template directory (cover.docx, table.docx, header-footer.docx, etc.)
//! - Apply templates to the generated document
//! - Discover markdown files using config patterns
//! - Generate a DOCX file with all settings from config and templates
//!
//! Run with: cargo run --example generate_thai_manual

use md2docx::config::ProjectConfig;
use md2docx::discovery::DiscoveredProject;
use md2docx::docx::ooxml::{FooterConfig, HeaderConfig, HeaderFooterField};
use md2docx::docx::TocConfig;
use md2docx::{
    markdown_to_docx_with_templates, DocumentConfig, DocumentMeta, Language, PlaceholderContext,
    TemplateDir,
};
use std::fs;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("📚 กำลังสร้างคู่มือภาษาไทย...");
    println!("   Generating Thai Manual...\n");

    let base_dir = Path::new("examples/thai-manual");

    // Load config from md2docx.toml
    let config_path = base_dir.join("md2docx.toml");
    let project_config = if config_path.exists() {
        println!("📖 Loading config from md2docx.toml");
        let config_str = fs::read_to_string(&config_path)?;
        #[cfg(feature = "cli")]
        let config = ProjectConfig::parse_toml(&config_str)?;
        #[cfg(not(feature = "cli"))]
        let config = toml::from_str::<ProjectConfig>(&config_str)?;
        config
    } else {
        println!("⚠️  No md2docx.toml found, using defaults");
        ProjectConfig::default()
    };

    // Load templates from template directory if configured
    let (template_set, template_loaded, header_footer_template) =
        if let Some(template_dir) = project_config.template.dir.as_ref() {
            let template_path = base_dir.join(template_dir);
            if template_path.exists() {
                println!("🎨 Loading templates from: {:?}", template_path);
                let template_dir_obj = TemplateDir::load(&template_path)?;
                let templates = template_dir_obj.load_all()?;

                if templates.has_cover() {
                    println!("   ✓ Cover template found (cover.docx)");
                }
                if templates.has_table() {
                    println!("   ✓ Table template found (table.docx)");
                }
                if templates.has_image() {
                    println!("   ✓ Image template found (image.docx)");
                }
                if templates.has_header_footer() {
                    println!("   ✓ Header/Footer template found (header-footer.docx)");
                }

                // Load header/footer template separately for DocumentConfig
                let hf_template = {
                    let hf_path = template_path.join("header-footer.docx");
                    if hf_path.exists() {
                        println!("   Loading header/footer template...");
                        match md2docx::template::extract::header_footer::extract(&hf_path) {
                            Ok(template) => {
                                println!(
                                    "     Different first page: {}",
                                    template.different_first_page
                                );
                                Some(template)
                            }
                            Err(e) => {
                                eprintln!("   ⚠️  Failed to load header-footer.docx: {}", e);
                                None
                            }
                        }
                    } else {
                        None
                    }
                };

                (Some(templates), true, hf_template)
            } else {
                println!("⚠️  Template directory not found: {:?}", template_path);
                (None, false, None)
            }
        } else {
            (None, false, None)
        };

    // Discover files using config patterns
    println!("\n🔍 Discovering files in: {:?}", base_dir);
    let project = DiscoveredProject::discover_with_config(base_dir, &project_config)?;

    if !project.is_valid() {
        return Err(
            "No markdown files found. Please check your md2docx.toml configuration.".into(),
        );
    }

    // Get ordered list of files
    let files = project.all_files();

    println!("📁 Found {} files to process:", files.len());
    for file in &files {
        println!(
            "   - {}",
            file.file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("unknown")
        );
    }

    // Extract inside content from cover.md BEFORE we potentially skip it
    // This will be used for {{inside}} placeholder in cover template
    let _inside_content = extract_cover_inside_content(base_dir);

    // Read and combine all files (excluding cover.md if using cover template)
    let mut combined_markdown = String::new();
    let using_cover_template = template_set
        .as_ref()
        .map(|t| t.has_cover())
        .unwrap_or(false);

    // Track the directory of the first non-cover markdown file for image resolution
    let mut first_content_file_dir: Option<std::path::PathBuf> = None;

    for (_i, file_path) in files.iter().enumerate() {
        let file_name = file_path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("unknown");

        // Skip cover.md if using cover template (it's rendered via {{inside}} placeholder)
        if using_cover_template && file_name == "cover.md" {
            println!(
                "📖 Reading: {} (using cover template, skipping from main content)",
                file_name
            );
            continue;
        }

        println!("📖 Reading: {}", file_name);

        // Track the directory of the first content file for image resolution
        if first_content_file_dir.is_none() {
            if let Some(parent) = file_path.parent() {
                first_content_file_dir = Some(parent.to_path_buf());
            }
        }

        let content = fs::read_to_string(file_path)?;

        // Strip frontmatter from each file
        let content_without_frontmatter = strip_frontmatter(&content);

        // Add page break between chapters (except before first)
        if !combined_markdown.is_empty() {
            combined_markdown.push_str("\n\n---\n\n");
        }

        combined_markdown.push_str(&content_without_frontmatter);
    }

    // Determine language from config
    let lang = if project_config.is_thai() {
        Language::Thai
    } else {
        Language::English
    };

    // Build header config from project config
    // ProjectConfig no longer has explicit header/footer sections in the schema
    // so we use defaults or build from other config if needed.
    // For this example, we'll use defaults.
    let header_config = HeaderConfig::default();
    let footer_config = FooterConfig::default();

    // Map FontsSection to FontConfig
    // Note: Font sizes are in half-points (multiply pt by 2 for OOXML)
    let font_config = Some(md2docx::docx::ooxml::FontConfig {
        default: if project_config.fonts.default.is_empty() {
            None
        } else {
            Some(project_config.fonts.default.clone())
        },
        code: if project_config.fonts.code.is_empty() {
            None
        } else {
            Some(project_config.fonts.code.clone())
        },
        normal_size: Some(project_config.fonts.normal_based_size * 2),
        normal_color: Some(project_config.fonts.normal_based_color.clone()),
        h1_color: Some(project_config.fonts.h1_based_color.clone()),
        caption_size: Some(project_config.fonts.caption_based_size * 2),
        caption_color: Some(project_config.fonts.caption_based_color.clone()),
        code_size: Some(project_config.fonts.code_based_size * 2),
    });

    // Build document config from project config
    let doc_config = DocumentConfig {
        title: project_config.document.title.clone(),
        toc: TocConfig {
            enabled: project_config.toc.enabled,
            depth: project_config.toc.depth,
            title: project_config.toc.title.clone(),
            // If using template, we handle cover insertion separately (prepend mode),
            // so we set after_cover = false so TOC is generated at the top of the body
            // (and then pushed down by the cover).
            after_cover: if template_loaded {
                false
            } else {
                project_config.toc.after_cover
            },
        },
        header: header_config,
        footer: footer_config,
        different_first_page: false, // Default to false as page_numbers section was removed
        template_dir: project_config
            .template
            .dir
            .as_ref()
            .map(|d| base_dir.join(d)),
        id_offset: 0,
        // When using templates, cover is handled separately, so include all headings in TOC
        process_all_headings: template_loaded,
        header_footer_template,
        document_meta: Some(DocumentMeta {
            title: project_config.document.title.clone(),
            subtitle: project_config.document.subtitle.clone(),
            author: project_config.document.author.clone(),
            date: project_config.date(),
        }),
        fonts: font_config,
        // Set base path for resolving relative image paths
        base_path: first_content_file_dir,
        page: None, // TODO: Add page config support from project config
    };

    // Extract inside content from cover.md (content after frontmatter)
    let inside_content = extract_cover_inside_content(base_dir);

    // Create placeholder context for templates
    let mut placeholder_ctx = PlaceholderContext {
        title: project_config.document.title.clone(),
        subtitle: project_config.document.subtitle.clone(),
        author: project_config.document.author.clone(),
        date: project_config.date(),
        version: "".to_string(),
        chapter: "".to_string(),
        page: "".to_string(),
        total: "".to_string(),
        custom: std::collections::HashMap::new(),
    };

    // Add inside content from cover.md
    if let Some(inside) = inside_content {
        placeholder_ctx = placeholder_ctx.with_custom("inside", inside);
    }

    println!("\n✨ Generating DOCX...");
    println!("   Language: {:?}", lang);
    println!("   TOC enabled: {}", doc_config.toc.enabled);
    println!("   TOC title: {}", doc_config.toc.title);
    if template_loaded {
        println!("   Templates: Enabled");
    }

    // Generate DOCX with templates
    let docx_bytes = if template_loaded {
        // Use template-aware generation
        markdown_to_docx_with_templates(
            &combined_markdown,
            lang,
            &doc_config,
            template_set.as_ref(),
            &placeholder_ctx,
        )?
    } else {
        // Fall back to standard generation
        use md2docx::markdown_to_docx_with_config;
        markdown_to_docx_with_config(&combined_markdown, lang, &doc_config)?
    };

    // Determine output path from config or use default
    // Use resolve_filename to expand placeholders like {{currenttime:...}} and {{title}}
    let output_path = if let Some(resolved) = project_config
        .output
        .resolve_filename(Some(&project_config))
    {
        base_dir.join(resolved)
    } else {
        base_dir.join("output/manual.docx")
    };

    // Create output directory if needed
    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)?;
    }

    fs::write(&output_path, &docx_bytes)?;

    println!("✅ Complete!");
    println!("📄 File: {:?}", output_path);
    println!("📊 Size: {} KB", docx_bytes.len() / 1024);

    if template_loaded {
        println!("\n🎨 Template features applied:");
        if template_set
            .as_ref()
            .map(|t| t.has_cover())
            .unwrap_or(false)
        {
            println!("   • Cover page with placeholders");
        }
        if template_set
            .as_ref()
            .map(|t| t.has_table())
            .unwrap_or(false)
        {
            println!("   • Table styles (header, odd/even rows)");
        }
        if template_set
            .as_ref()
            .map(|t| t.has_image())
            .unwrap_or(false)
        {
            println!("   • Image caption styles");
        }
        if template_set
            .as_ref()
            .map(|t| t.has_header_footer())
            .unwrap_or(false)
        {
            println!("   • Header/footer styles");
        }
    }

    Ok(())
}

/// Parse a header/footer field string from config into HeaderFooterField variants
fn parse_header_footer_field(s: &str) -> Vec<HeaderFooterField> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return vec![];
    }

    match trimmed {
        "{title}" => vec![HeaderFooterField::DocumentTitle],
        "{page}" => vec![HeaderFooterField::PageNumber],
        "{total}" => vec![HeaderFooterField::TotalPages],
        "{chapter}" => vec![HeaderFooterField::ChapterName],
        other => vec![HeaderFooterField::Text(other.to_string())],
    }
}

/// Strip YAML frontmatter from markdown content
fn strip_frontmatter(content: &str) -> String {
    if !content.starts_with("---") {
        return content.to_string();
    }

    let lines: Vec<&str> = content.lines().collect();
    let mut closing_line = None;

    for (i, line) in lines.iter().enumerate().skip(1) {
        if line.trim() == "---" {
            closing_line = Some(i);
            break;
        }
    }

    match closing_line {
        Some(idx) => lines[idx + 1..].join("\n"),
        None => content.to_string(),
    }
}

/// Extract inside content from cover.md for {{inside}} placeholder
fn extract_cover_inside_content(base_dir: &Path) -> Option<String> {
    let cover_path = base_dir.join("cover.md");
    if !cover_path.exists() {
        return None;
    }

    let content = fs::read_to_string(cover_path).ok()?;

    // Extract content after frontmatter (the "inside" content)
    let inside = strip_frontmatter(&content);

    // Trim whitespace but keep the content
    let trimmed = inside.trim();

    if trimmed.is_empty() {
        return None;
    }

    // Fix image paths to be relative to project root
    // The cover.md is at base_dir, so paths like "assets/logo.png" need to be prefixed
    // if base_dir is not "."
    let fixed_content = if base_dir.components().count() > 0 {
        let prefix = format!("{}/", base_dir.display());
        trimmed.replace("assets/", &format!("{}assets/", prefix))
    } else {
        trimmed.to_string()
    };

    Some(fixed_content)
}
