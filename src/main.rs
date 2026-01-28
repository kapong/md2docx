//! md2docx CLI entry point

#[cfg(feature = "cli")]
use clap::{Parser, Subcommand};
use md2docx::docx::ooxml::HeaderFooterField;
use std::path::PathBuf;

#[cfg(feature = "cli")]
#[derive(Parser)]
#[command(name = "md2docx")]
#[command(author, version, about = "Convert Markdown to DOCX", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[cfg(feature = "cli")]
#[derive(Subcommand)]
enum Commands {
    /// Build DOCX from markdown files
    Build {
        /// Input markdown file
        #[arg(short, long)]
        input: Option<PathBuf>,

        /// Input directory with chapter files
        #[arg(short, long)]
        dir: Option<PathBuf>,

        /// Output DOCX file
        #[arg(short, long)]
        output: PathBuf,

        /// Template DOCX file
        #[arg(long)]
        template: Option<PathBuf>,

        /// Include table of contents
        #[arg(long)]
        toc: bool,
    },

    /// Generate a template DOCX with all styles for customization
    DumpTemplate {
        /// Output file path
        #[arg(short, long, default_value = "template.docx")]
        output: PathBuf,

        /// Language for default fonts (en or th)
        #[arg(long, default_value = "en")]
        lang: String,

        /// Generate minimal template (fewer styles)
        #[arg(long)]
        minimal: bool,
    },

    /// Validate a template has all required styles
    ValidateTemplate {
        /// Template file to validate
        template: PathBuf,
    },
}

#[cfg(feature = "cli")]
/// Parse a header/footer field string from config into HeaderFooterField variants
///
/// Supported patterns:
/// - "{title}" -> HeaderFooterField::DocumentTitle
/// - "{page}" -> HeaderFooterField::PageNumber
/// - "{total}" -> HeaderFooterField::TotalPages
/// - "{chapter}" -> HeaderFooterField::ChapterName
/// - "" (empty) -> empty vec
/// - "any other text" -> HeaderFooterField::Text(text)
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

#[cfg(feature = "cli")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Build {
            input,
            dir,
            output,
            template,
            toc,
        } => {
            use md2docx::config::ProjectConfig;
            use md2docx::discovery::DiscoveredProject;
            use md2docx::docx::ooxml::{FooterConfig, HeaderConfig};
            use md2docx::{markdown_to_docx_with_config, DocumentConfig, Language};

            // Ignore template option for now (not yet implemented)
            if let Some(ref tmpl) = template {
                eprintln!(
                    "Warning: Template support is not yet implemented. Ignoring --template {:?}",
                    tmpl
                );
            }

            // Determine source mode: single file or directory
            // Store the base directory for resolving relative image paths
            let (content, lang, doc_config, base_dir) = if let Some(input_file) = input {
                // Single file mode
                println!("Reading input file: {}", input_file.display());
                let markdown = std::fs::read_to_string(&input_file)?;
                // Set base directory to the input file's parent directory
                let base_dir = input_file.parent().map(|p| p.to_path_buf());
                let lang = Language::English; // Default
                let mut config = DocumentConfig::default();
                if toc {
                    config.toc.enabled = true;
                }
                (markdown, lang, config, base_dir)
            } else if let Some(input_dir) = dir {
                // Directory mode
                println!("Scanning directory: {}", input_dir.display());

                // a. Load config if exists
                let config_path = input_dir.join("md2docx.toml");
                let project_config: ProjectConfig = if config_path.exists() {
                    println!("Loading config from: {}", config_path.display());
                    ProjectConfig::from_file(&config_path)?
                } else {
                    println!("No md2docx.toml found, using defaults");
                    ProjectConfig::default()
                };

                // b. Discover files
                let project = DiscoveredProject::discover_with_config(&input_dir, &project_config)?;

                if !project.is_valid() {
                    eprintln!("Error: No markdown files found in {:?}", input_dir);
                    std::process::exit(1);
                }

                println!("Found {} chapter(s)", project.chapters.len());
                if project.cover.is_some() {
                    println!("Found cover page");
                }
                if !project.appendices.is_empty() {
                    println!("Found {} appendix/appendices", project.appendices.len());
                }
                if project.bibliography.is_some() {
                    println!("Found bibliography");
                }

                // c. Merge all files into single markdown
                let mut combined = String::new();
                for file_path in project.all_files() {
                    println!(
                        "  Including: {}",
                        file_path.file_name().unwrap_or_default().to_string_lossy()
                    );
                    let content = std::fs::read_to_string(file_path)?;
                    if !combined.is_empty() {
                        // Add section break between files
                        combined.push_str("\n\n---\n\n");
                    }
                    combined.push_str(&content);
                }

                // d. Determine language
                let lang = if project_config.is_thai() {
                    Language::Thai
                } else {
                    Language::English
                };

                // e. Build DocumentConfig from ProjectConfig
                // Parse header fields - use defaults if all fields are empty
                let header_config = if project_config.header.left.is_empty()
                    && project_config.header.center.is_empty()
                    && project_config.header.right.is_empty()
                {
                    // Use default header config (DocumentTitle on left, ChapterName on right)
                    HeaderConfig::default()
                } else {
                    HeaderConfig {
                        left: parse_header_footer_field(&project_config.header.left),
                        center: parse_header_footer_field(&project_config.header.center),
                        right: parse_header_footer_field(&project_config.header.right),
                    }
                };

                // Parse footer fields - use defaults if all fields are empty
                let footer_config = if project_config.footer.left.is_empty()
                    && project_config.footer.center.is_empty()
                    && project_config.footer.right.is_empty()
                {
                    // Use default footer config (PageNumber in center)
                    FooterConfig::default()
                } else {
                    FooterConfig {
                        left: parse_header_footer_field(&project_config.footer.left),
                        center: parse_header_footer_field(&project_config.footer.center),
                        right: parse_header_footer_field(&project_config.footer.right),
                    }
                };

                let doc_config = DocumentConfig {
                    title: project_config.document.title.clone(),
                    toc: md2docx::docx::TocConfig {
                        enabled: project_config.toc.enabled || toc, // CLI override
                        depth: project_config.toc.depth,
                        title: project_config.toc.title.clone(),
                        ..Default::default()
                    },
                    header: header_config,
                    footer: footer_config,
                    different_first_page: project_config.page_numbers.skip_chapter_first,
                    template_dir: None,
                    id_offset: 0,
                    process_all_headings: false,
                };

                (combined, lang, doc_config, Some(input_dir.clone()))
            } else {
                eprintln!("Error: Either --input or --dir must be specified");
                std::process::exit(1);
            };

            // 2. Convert to DOCX
            println!("Converting to DOCX...");

            // Change to base directory so relative image paths work correctly
            let original_dir = std::env::current_dir()?;
            if let Some(ref base) = base_dir {
                std::env::set_current_dir(base)?;
            }

            let docx_bytes = markdown_to_docx_with_config(&content, lang, &doc_config)?;

            // Restore original directory
            std::env::set_current_dir(original_dir)?;

            // 3. Write output
            std::fs::write(&output, docx_bytes)?;
            println!("Successfully created: {}", output.display());
        }

        Commands::DumpTemplate {
            output,
            lang,
            minimal,
        } => {
            use md2docx::docx::ooxml::Language;
            use md2docx::docx::template::generate_template;

            let language = match lang.to_lowercase().as_str() {
                "th" | "thai" => Language::Thai,
                _ => Language::English,
            };

            println!(
                "Generating template with {} defaults...",
                if language == Language::Thai {
                    "Thai"
                } else {
                    "English"
                }
            );

            let bytes = generate_template(language, minimal)?;
            std::fs::write(&output, bytes)?;

            println!("Template written to: {}", output.display());
            println!("\nNext steps:");
            println!("1. Open {} in Microsoft Word", output.display());
            println!("2. Modify styles (Home → Styles → right-click → Modify)");
            println!(
                "3. Save and use with: md2docx build --template {}",
                output.display()
            );
        }

        Commands::ValidateTemplate { template } => {
            use md2docx::docx::template::validate_template;

            if !template.exists() {
                eprintln!("Error: Template file not found: {:?}", template);
                std::process::exit(1);
            }

            println!("Validating template: {}", template.display());

            let result = validate_template(&template)?;

            println!();
            if result.is_valid() {
                println!("✓ Template is valid! All required styles are present.");
            } else {
                println!("✗ Template is missing required styles:");
                for style in &result.missing_required {
                    println!("  - {}", style);
                }
            }

            if !result.missing_recommended.is_empty() {
                println!();
                println!("Optional styles not found (these will use defaults):");
                for style in &result.missing_recommended {
                    println!("  - {}", style);
                }
            }

            if !result.warnings.is_empty() {
                println!();
                println!("Warnings:");
                for warning in &result.warnings {
                    println!("  ! {}", warning);
                }
            }

            println!();
            println!("Found {} styles in template", result.found_styles.len());

            if !result.is_valid() {
                println!();
                println!("Tip: Use 'md2docx dump-template -o template.docx' to generate a valid template.");
                std::process::exit(1);
            }
        }
    }

    Ok(())
}

#[cfg(not(feature = "cli"))]
fn main() {
    compile_error!("CLI feature is required for the binary");
}
