//! Integration test for table template extraction
//!
//! This test verifies that table template extraction works with a real DOCX file.

use md2docx::template::extract::table::extract;
use std::path::Path;

#[test]
fn test_extract_from_real_table_docx() {
    let table_path = Path::new("examples/thai-manual/template/table.docx");

    // Skip test if file doesn't exist (might not be in all environments)
    if !table_path.exists() {
        println!(
            "Skipping test: table.docx not found at {}",
            table_path.display()
        );
        return;
    }

    let result = extract(table_path);

    match result {
        Ok(template) => {
            println!("Successfully extracted table template");
            println!("Header background: {:?}", template.header.background_color);
            println!("Header font: {}", template.header.font_family);
            println!(
                "Header font size: {} half-points",
                template.header.font_size
            );
            println!("Header font color: {}", template.header.font_color);
            println!("Header bold: {}", template.header.bold);

            println!(
                "Odd row background: {:?}",
                template.row_odd.background_color
            );
            println!(
                "Even row background: {:?}",
                template.row_even.background_color
            );

            println!("First column bold: {}", template.first_column.bold);
            println!("Other columns bold: {}", template.other_columns.bold);

            println!("Border top style: {}", template.borders.top.style);
            println!("Border top color: {}", template.borders.top.color);
            println!("Border top width: {}", template.borders.top.width);

            println!("Caption font: {}", template.caption.font_family);
            println!(
                "Caption font size: {} half-points",
                template.caption.font_size
            );
            println!("Caption italic: {}", template.caption.italic);

            // Verify we got some reasonable values
            assert!(!template.header.font_family.is_empty());
            assert!(template.header.font_size > 0);
            assert!(!template.header.font_color.is_empty());

            // Border should have some style
            assert!(!template.borders.top.style.is_empty());
            assert!(!template.borders.top.color.is_empty());
        }
        Err(e) => {
            panic!("Failed to extract table template: {}", e);
        }
    }
}
