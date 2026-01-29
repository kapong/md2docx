//! md2docx CLI entry point

#[cfg(feature = "cli")]
use clap::{Parser, Subcommand};
use regex::Regex;
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
        output: Option<PathBuf>,

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
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Build {
            input,
            dir,
            output,
            template: _,
            toc,
        } => {
            use md2docx::config::ProjectConfig;
            use md2docx::discovery::DiscoveredProject;
            use md2docx::{
                markdown_to_docx_with_templates, DocumentConfig, Language, PlaceholderContext,
            };

            // 1. Determine base directory and load config
            let (input_dir, project_config) = if let Some(ref input_dir) = dir {
                let config_path = input_dir.join("md2docx.toml");
                let config = if config_path.exists() {
                    println!("Loading config from: {}", config_path.display());
                    ProjectConfig::from_file(&config_path)?
                } else {
                    println!("No md2docx.toml found, using defaults");
                    ProjectConfig::default()
                };
                (input_dir.clone(), config)
            } else if let Some(ref input_file) = input {
                let input_dir = input_file
                    .parent()
                    .unwrap_or(std::path::Path::new("."))
                    .to_path_buf();

                // Try to load config from input directory
                let config_path = input_dir.join("md2docx.toml");
                let config = if config_path.exists() {
                    println!("Loading config from: {}", config_path.display());
                    ProjectConfig::from_file(&config_path)?
                } else {
                    ProjectConfig::default()
                };
                (input_dir, config)
            } else {
                eprintln!("Error: Either --input or --dir must be specified");
                std::process::exit(1);
            };

            // 2. Load templates from template.dir
            let template_set = if let Some(ref template_dir_path) = project_config.template.dir {
                let template_path = input_dir.join(template_dir_path);
                if template_path.exists() {
                    println!("Loading templates from: {}", template_path.display());
                    match md2docx::TemplateDir::load(&template_path) {
                        Ok(template_dir) => match template_dir.load_all() {
                            Ok(set) => {
                                if set.has_cover() {
                                    println!("  Found cover.docx");
                                }
                                if set.has_table() {
                                    println!("  Found table.docx");
                                }
                                if set.has_image() {
                                    println!("  Found image.docx");
                                }
                                if set.has_header_footer() {
                                    println!("  Found header-footer.docx");
                                }
                                Some(set)
                            }
                            Err(e) => {
                                eprintln!("Warning: Failed to load templates: {}", e);
                                None
                            }
                        },
                        Err(e) => {
                            eprintln!("Warning: Failed to open template directory: {}", e);
                            None
                        }
                    }
                } else {
                    eprintln!(
                        "Warning: Template directory not found: {}",
                        template_path.display()
                    );
                    None
                }
            } else {
                None
            };

            // 3. Discover files and merge markdown
            let content = if let Some(ref input_file) = input {
                println!("Reading input file: {}", input_file.display());
                std::fs::read_to_string(input_file)?
            } else {
                let project = DiscoveredProject::discover_with_config(&input_dir, &project_config)?;
                if !project.is_valid() {
                    eprintln!("Error: No markdown files found in {:?}", input_dir);
                    std::process::exit(1);
                }
                println!("Found {} chapter(s)", project.chapters.len());
                let mut combined = String::new();

                for file_path in project.all_files() {
                    println!(
                        "  Including: {}",
                        file_path.file_name().unwrap_or_default().to_string_lossy()
                    );
                    let raw_content = std::fs::read_to_string(file_path)?;

                    // Rewrite relative image paths to be relative to CWD
                    let file_content = resolve_image_paths(&raw_content, file_path);

                    if !combined.is_empty() {
                        combined.push_str("\n\n---\n\n");
                    }
                    combined.push_str(&file_content);
                }
                combined
            };

            // 4. Create placeholder context from config
            let mut placeholder_ctx = PlaceholderContext::default();
            placeholder_ctx.set("title", &project_config.document.title);
            placeholder_ctx.set("subtitle", &project_config.document.subtitle);
            placeholder_ctx.set("author", &project_config.document.author);
            placeholder_ctx.set("date", &project_config.date());
            placeholder_ctx.set("version", &project_config.document.version);

            // 5. Determine language
            let lang = if project_config.is_thai() {
                Language::Thai
            } else {
                Language::English
            };

            // 6. Build document config
            let doc_config = DocumentConfig {
                title: project_config.document.title.clone(),
                toc: md2docx::docx::TocConfig {
                    enabled: project_config.toc.enabled || toc,
                    depth: project_config.toc.depth,
                    title: project_config.toc.title.clone(),
                    after_cover: project_config.toc.after_cover,
                },
                header_footer_template: template_set.as_ref().and_then(|t| t.header_footer.clone()),
                document_meta: Some(md2docx::DocumentMeta {
                    title: project_config.document.title.clone(),
                    subtitle: project_config.document.subtitle.clone(),
                    author: project_config.document.author.clone(),
                    date: project_config.date(),
                }),
                fonts: Some(md2docx::docx::FontConfig {
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
                }),
                ..DocumentConfig::default()
            };

            // 7. Convert with templates
            println!("Converting to DOCX...");
            let original_dir = std::env::current_dir()?;
            std::env::set_current_dir(&input_dir)?;
            let docx_bytes = markdown_to_docx_with_templates(
                &content,
                lang,
                &doc_config,
                template_set.as_ref(),
                &placeholder_ctx,
            )?;
            std::env::set_current_dir(original_dir)?;

            // 8. Determine final output path
            let final_output = if let Some(ref cli_output) = output {
                cli_output.clone()
            } else if let Some(resolved) = project_config
                .output
                .resolve_filename(Some(&project_config))
            {
                input_dir.join(resolved)
            } else {
                eprintln!(
                    "Error: No output file specified. Use --output or set output.file in config."
                );
                std::process::exit(1);
            };

            // 9. Create parent dirs and write
            if let Some(parent) = final_output.parent() {
                if !parent.exists() {
                    std::fs::create_dir_all(parent)?;
                }
            }
            std::fs::write(&final_output, docx_bytes)?;
            println!("Successfully created: {}", final_output.display());
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

/// Rewrite image paths in markdown content to be relative to the markdown file's directory.
fn resolve_image_paths(content: &str, file_path: &std::path::Path) -> String {
    if let Some(parent) = file_path.parent() {
        // Regex for rewriting image paths: ![alt](url "title") - handling optional whitespace
        let image_regex = Regex::new(r"!\[(.*?)\]\s*\((.*?)\)").expect("Invalid regex");

        image_regex
            .replace_all(content, |caps: &regex::Captures| {
                let alt = &caps[1];
                let raw_link = &caps[2];

                // Trim leading/trailing whitespace from the link content
                let link_content = raw_link.trim();

                // Split url and optional title
                let (url, title_suffix) = match link_content.find(char::is_whitespace) {
                    Some(idx) => (&link_content[..idx], &link_content[idx..]),
                    None => (link_content, ""),
                };

                // Skip absolute URLs, absolute paths, or data URIs
                if url.starts_with("http://")
                    || url.starts_with("https://")
                    || url.starts_with("/")
                    || url.starts_with("data:")
                    || std::path::Path::new(url).is_absolute()
                {
                    return caps[0].to_string();
                }

                // Resolve relative to file parent
                let new_path = parent.join(url);
                // Normalize to forward slashes for consistency
                let new_path_str = new_path.to_string_lossy().replace('\\', "/");

                format!("![{}]({}{})", alt, new_path_str, title_suffix)
            })
            .to_string()
    } else {
        content.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    #[test]
    fn test_resolve_image_paths_relative() {
        let content = "![Image](img.png)";
        let file_path = Path::new("docs/chapter1.md");
        let result = resolve_image_paths(content, file_path);
        assert_eq!(result, "![Image](docs/img.png)");
    }

    #[test]
    fn test_resolve_image_paths_with_title() {
        let content = "![Image](img.png \"My Title\")";
        let file_path = Path::new("docs/chapter1.md");
        let result = resolve_image_paths(content, file_path);
        assert_eq!(result, "![Image](docs/img.png \"My Title\")");
    }

    #[test]
    fn test_resolve_image_paths_absolute_url() {
        let content = "![Image](https://example.com/img.png)";
        let file_path = Path::new("docs/chapter1.md");
        let result = resolve_image_paths(content, file_path);
        assert_eq!(result, content);
    }

    #[test]
    fn test_resolve_image_paths_whitespace() {
        let content = "![Image](  img.png  )";
        let file_path = Path::new("docs/ch1.md");
        let result = resolve_image_paths(content, file_path);
        assert_eq!(result, "![Image](docs/img.png)");
    }
}

#[cfg(not(feature = "cli"))]
fn main() {
    compile_error!("CLI feature is required for the binary");
}
