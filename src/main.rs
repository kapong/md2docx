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
