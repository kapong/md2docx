//! Example: Generate Thai Manual from Markdown files
//!
//! This example demonstrates how to:
//! - Read multiple markdown files
//! - Load and parse a config file (md2docx.toml)
//! - Combine files into a single document
//! - Generate a DOCX file with proper TOC, headers, and footers from config
//!
//! Run with: cargo run --example generate_thai_manual

use md2docx::config::ProjectConfig;
use md2docx::docx::ooxml::{FooterConfig, HeaderConfig, HeaderFooterField};
use md2docx::docx::TocConfig;
use md2docx::{markdown_to_docx_with_config, DocumentConfig, Language};
use std::fs;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("📚 กำลังสร้างคู่มือภาษาไทย...");
    println!("   Generating Thai Manual...\n");

    let base_dir = Path::new("examples/thai-manual");
    let output_dir = base_dir.join("output");

    // Create output directory if it doesn't exist
    fs::create_dir_all(&output_dir)?;

    // Load config from md2docx.toml
    let config_path = base_dir.join("md2docx.toml");
    let project_config = if config_path.exists() {
        println!("📖 Loading config from md2docx.toml");
        let config_str = fs::read_to_string(&config_path)?;
        toml::from_str::<ProjectConfig>(&config_str)?
    } else {
        println!("⚠️  No md2docx.toml found, using defaults");
        ProjectConfig::default()
    };

    // Define chapters in order
    let chapters = [
        "cover.md",
        "ch01_introduction.md",
        "ch02_markdown_syntax.md",
        "ch03_configuration.md",
        "ch04_advanced.md",
    ];

    // Read and combine all chapters
    let mut combined_markdown = String::new();

    for (i, chapter) in chapters.iter().enumerate() {
        let path = base_dir.join(chapter);
        println!("📖 อ่านไฟล์: {}", chapter);

        let content = fs::read_to_string(&path)?;

        // Strip frontmatter from each chapter
        let content_without_frontmatter = strip_frontmatter(&content);

        // Fix image paths: replace "assets/" with "examples/thai-manual/assets/"
        // This is needed because the tool runs from project root, but files are in subdirectory
        let content_with_fixed_paths =
            content_without_frontmatter.replace("assets/", "examples/thai-manual/assets/");

        // Add page break between chapters (except before cover)
        if i > 0 {
            combined_markdown.push_str("\n\n---\n\n"); // Page break
        }

        combined_markdown.push_str(&content_with_fixed_paths);
    }

    // Determine language
    let lang = if project_config.is_thai() {
        Language::Thai
    } else {
        Language::English
    };

    // Build header config - use defaults if all empty
    let header_config = if project_config.header.left.is_empty()
        && project_config.header.center.is_empty()
        && project_config.header.right.is_empty()
    {
        HeaderConfig::default()
    } else {
        HeaderConfig {
            left: parse_header_footer_field(&project_config.header.left),
            center: parse_header_footer_field(&project_config.header.center),
            right: parse_header_footer_field(&project_config.header.right),
        }
    };

    // Build footer config - use defaults if all empty
    let footer_config = if project_config.footer.left.is_empty()
        && project_config.footer.center.is_empty()
        && project_config.footer.right.is_empty()
    {
        FooterConfig::default()
    } else {
        FooterConfig {
            left: parse_header_footer_field(&project_config.footer.left),
            center: parse_header_footer_field(&project_config.footer.center),
            right: parse_header_footer_field(&project_config.footer.right),
        }
    };

    // Build document config
    let doc_config = DocumentConfig {
        title: project_config.document.title.clone(),
        toc: TocConfig {
            enabled: project_config.toc.enabled,
            depth: project_config.toc.depth,
            title: project_config.toc.title.clone(),
            after_cover: true,
        },
        header: header_config,
        footer: footer_config,
        different_first_page: project_config.page_numbers.skip_chapter_first,
    };

    println!("\n✨ กำลังสร้างไฟล์ DOCX...");
    println!("   Language: {:?}", lang);
    println!("   TOC enabled: {}", doc_config.toc.enabled);
    println!("   TOC title: {}", doc_config.toc.title);

    // Generate DOCX with config
    let docx_bytes = markdown_to_docx_with_config(&combined_markdown, lang, &doc_config)?;

    // Save to file
    let output_path = output_dir.join("manual-md2docx.docx");
    fs::write(&output_path, &docx_bytes)?;

    println!("✅ สร้างเสร็จแล้ว!");
    println!("📄 ไฟล์: {:?}", output_path);
    println!("📊 ขนาด: {} KB", docx_bytes.len() / 1024);

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
    // Check if content starts with ---
    if !content.starts_with("---") {
        return content.to_string();
    }

    let lines: Vec<&str> = content.lines().collect();

    // Find closing ---
    let mut closing_line = None;
    for (i, line) in lines.iter().enumerate().skip(1) {
        if line.trim() == "---" {
            closing_line = Some(i);
            break;
        }
    }

    match closing_line {
        Some(idx) => {
            // Return everything after the closing ---
            lines[idx + 1..].join("\n")
        }
        None => content.to_string(),
    }
}
