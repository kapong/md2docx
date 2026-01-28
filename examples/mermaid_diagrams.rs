//! Example: Mermaid Diagram Rendering
//!
//! This example demonstrates rendering various Mermaid diagram types
//! to DOCX using the native Rust mermaid-rs-renderer library.
//!
//! Run with: cargo run --example mermaid_diagrams

use md2docx::{markdown_to_docx_with_config, DocumentConfig, Language};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🧜‍♀️ Mermaid Diagram Rendering Example\n");

    // Read the markdown file
    let markdown = fs::read_to_string("examples/mermaid-demo.md")?;

    println!("📄 Input file: examples/mermaid-demo.md");

    println!("📊 Markdown size: {} bytes\n", markdown.len());

    // Convert to DOCX
    println!("⚙️  Rendering diagrams...");
    let start = std::time::Instant::now();

    let docx_bytes =
        markdown_to_docx_with_config(&markdown, Language::English, &DocumentConfig::default())?;

    let elapsed = start.elapsed();
    println!("✅ Rendered in {:.2?}\n", elapsed);

    // Save output
    let output_path = "mermaid_output.docx";
    fs::write(output_path, &docx_bytes)?;

    println!("📄 Output file: {}", output_path);
    println!("📦 File size: {} KB", docx_bytes.len() / 1024);

    println!(
        "\n🎉 Success! Open '{}' in Microsoft Word to see the rendered diagrams.",
        output_path
    );

    Ok(())
}
