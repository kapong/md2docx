// Example: Basic usage of md2docx library API

use md2docx::{Document, Language};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create a simple document
    let doc = Document::new()
        .add_heading(1, "My Document")
        .add_paragraph("This is a simple document created with md2docx.")
        .add_heading(2, "Introduction")
        .add_paragraph("Welcome to the md2docx library!")
        .add_quote("This is a blockquote example.")
        .add_heading(3, "Code Example")
        .add_code_block("fn main() {\n    println!(\"Hello, world!\");\n}");

    // Write to file
    doc.write_to_file("example_output.docx")?;
    println!("Document created: example_output.docx");

    // Create a Thai document
    let thai_doc = Document::with_language(Language::Thai)
        .add_heading(1, "เอกสารภาษาไทย")
        .add_paragraph("นี่คือเอกสารที่สร้างด้วย md2docx")
        .add_heading(2, "บทนำ")
        .add_paragraph("ยินดีต้อนรับสู่ไลบรารี md2docx!");

    thai_doc.write_to_file("example_thai.docx")?;
    println!("Thai document created: example_thai.docx");

    Ok(())
}
