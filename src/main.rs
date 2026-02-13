//! md2docx CLI entry point

#[cfg(feature = "cli")]
use clap::{Parser, Subcommand};
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
            use md2docx::project::ProjectBuilder;
            use md2docx::{
                markdown_to_docx_with_templates, DocumentConfig, Language, PlaceholderContext,
            };

            if let Some(ref input_dir) = dir {
                let mut builder = ProjectBuilder::from_directory(input_dir)?;

                // Apply CLI overrides
                if toc {
                    builder = builder.with_toc(true);
                }
                if let Some(ref out) = output {
                    builder = builder.with_output(out.clone());
                }

                // Build and write
                let output_path = builder.build_to_file()?;
                println!("Successfully created: {}", output_path.display());
            } else if let Some(ref input_file) = input {
                // Simple single file conversion
                println!("Reading input file: {}", input_file.display());
                let raw_content = std::fs::read_to_string(input_file)?;

                // Rewrite relative image paths
                let content = resolve_image_paths(&raw_content, input_file);

                // For single file, we use default config but can enable TOC if requested
                let mut doc_config = DocumentConfig::default();
                if toc {
                    doc_config.toc.enabled = true;
                }

                let docx_bytes = markdown_to_docx_with_templates(
                    &content,
                    Language::English,
                    &doc_config,
                    None,
                    &PlaceholderContext::default(),
                )?;

                let final_output = if let Some(ref out) = output {
                    out.clone()
                } else {
                    let mut out = input_file.clone();
                    out.set_extension("docx");
                    out
                };

                std::fs::write(&final_output, docx_bytes)?;
                println!("Successfully created: {}", final_output.display());
            } else {
                eprintln!("Error: Either --input or --dir must be specified");
                std::process::exit(1);
            }
        }
    }

    Ok(())
}

/// Rewrite image paths in markdown content to be relative to the markdown file's directory.
fn resolve_image_paths(content: &str, file_path: &std::path::Path) -> String {
    md2docx::project::resolve_image_paths(content, file_path)
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

    #[test]
    fn test_resolve_image_paths_skip_code_blocks() {
        let content = "![Outside](img.png)\n\n```markdown\n![Inside](assets/logo.png)\n```\n\n![After](other.png)";
        let file_path = Path::new("docs/chapter1.md");
        let result = resolve_image_paths(content, file_path);
        // Images outside code blocks should be resolved
        assert!(result.contains("![Outside](docs/img.png)"));
        assert!(result.contains("![After](docs/other.png)"));
        // Image inside code block should be preserved verbatim
        assert!(result.contains("![Inside](assets/logo.png)"));
    }
}

#[cfg(not(feature = "cli"))]
fn main() {
    compile_error!("CLI feature is required for the binary");
}
