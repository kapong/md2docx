//! Example: Using include resolution in md2docx
//!
//! This example demonstrates how to use the include resolution features
//! to pull content from external files into your DOCX documents.

use md2docx::{markdown_to_docx_with_includes, resolve_includes, IncludeConfig};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Setup: Create test files
    let temp_dir = std::env::temp_dir().join("md2docx_include_example");
    fs::create_dir_all(&temp_dir)?;

    // Create a markdown file to include
    let section_file = temp_dir.join("section.md");
    fs::write(
        &section_file,
        "## Included Section\n\nThis content was pulled from an external file!",
    )?;

    // Create a code file to include
    let code_file = temp_dir.join("example.rs");
    fs::write(
        &code_file,
        "fn main() {\n    println!(\"Hello from included code!\");\n}",
    )?;

    // Example 1: Using markdown_to_docx_with_includes (one-shot)
    println!("Example 1: One-shot conversion with includes");
    let markdown = r#"# Main Document

This is the main content.

{!include:section.md}

## Code Example

Here's some code from a source file:

{!code:example.rs}

End of document.
"#;

    let include_config = IncludeConfig {
        base_path: temp_dir.clone(),
        source_root: temp_dir.clone(),
        max_depth: 10,
    };

    let docx_bytes = markdown_to_docx_with_includes(markdown, &include_config)?;
    let output_path = temp_dir.join("output_with_includes.docx");
    fs::write(&output_path, docx_bytes)?;
    println!("  Created: {}", output_path.display());

    // Example 2: Using resolve_includes (manual control)
    println!("\nExample 2: Manual include resolution");
    use md2docx::parser::parse_markdown;

    let mut parsed = parse_markdown(markdown);
    resolve_includes(&mut parsed, &include_config)?;

    println!("  Resolved {} blocks", parsed.blocks.len());

    // Example 3: Nested includes
    println!("\nExample 3: Nested includes");
    let nested_file = temp_dir.join("nested.md");
    fs::write(
        &nested_file,
        "### Nested Content\n\nThis file was included by another included file.",
    )?;

    let parent_file = temp_dir.join("parent.md");
    fs::write(&parent_file, "## Parent Section\n\n{!include:nested.md}")?;

    let markdown_nested = r#"# Document with Nested Includes

{!include:parent.md}

End of document.
"#;

    let docx_bytes = markdown_to_docx_with_includes(markdown_nested, &include_config)?;
    let output_path_nested = temp_dir.join("output_nested.docx");
    fs::write(&output_path_nested, docx_bytes)?;
    println!("  Created: {}", output_path_nested.display());

    // Example 4: Code includes with line ranges
    println!("\nExample 4: Code includes with line ranges");
    let long_code_file = temp_dir.join("long_code.rs");
    fs::write(
        &long_code_file,
        "fn function1() {\n    println!(\"Line 2\");\n}\n\nfn function2() {\n    println!(\"Line 5\");\n}\n\nfn function3() {\n    println!(\"Line 8\");\n}\n",
    )?;

    let markdown_lines = r#"# Code with Line Ranges

Here's lines 2-5:

{!code:long_code.rs:2-5}

And here's just line 8:

{!code:long_code.rs:8-8}
"#;

    let docx_bytes = markdown_to_docx_with_includes(markdown_lines, &include_config)?;
    let output_path_lines = temp_dir.join("output_line_ranges.docx");
    fs::write(&output_path_lines, docx_bytes)?;
    println!("  Created: {}", output_path_lines.display());

    // Cleanup
    println!("\nCleaning up...");
    fs::remove_dir_all(temp_dir)?;

    println!("\nAll examples completed successfully!");
    Ok(())
}
