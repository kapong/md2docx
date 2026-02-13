//! Font embedding test — generates a DOCX with multiple embedded fonts
//!
//! Run: cargo run --example font_embedding

use md2docx::{markdown_to_docx_with_templates, DocumentConfig, Language, PlaceholderContext};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let font_dir = Path::new("docs/template/fonts");

    // Build a markdown document that uses multiple fonts via raw paragraphs
    let markdown = r#"# Font Embedding Test / ทดสอบการฝังฟอนต์

This document tests embedding multiple fonts into a DOCX file.

## Default Font (Pridi)

สวัสดีครับ นี่คือข้อความทดสอบด้วยฟอนต์ Pridi ซึ่งเป็นฟอนต์หลักของเอกสาร
The quick brown fox jumps over the lazy dog.

**ข้อความตัวหนา** — This should use Pridi Bold / ตัวหนาของ Pridi

## Thai Text Samples

### Sample 1: Standard paragraph
ภาษาไทยเป็นภาษาที่มีวรรณยุกต์ มีสระ พยัญชนะ และตัวเลขไทย ๐ ๑ ๒ ๓ ๔ ๕ ๖ ๗ ๘ ๙

### Sample 2: Mixed Thai-English
md2docx สามารถแปลง Markdown เป็นไฟล์ DOCX ได้อย่างง่ายดาย พร้อมรองรับภาษาไทยอย่างเต็มรูปแบบ

## Code Block (Consolas)

```rust
fn main() {
    println!("Hello, world! สวัสดีชาวโลก");
}
```

## Bullet List

- รายการที่ 1 — First item
- รายการที่ 2 — Second item  
- รายการที่ 3 — Third item

## Numbered List

1. ขั้นตอนแรก — First step
2. ขั้นตอนที่สอง — Second step
3. ขั้นตอนสุดท้าย — Final step

> นี่คือข้อความที่ยกมา / This is a blockquote
> ทดสอบภาษาไทยในบล็อกอ้างอิง

---

<!-- {font:Pridi} -->

This paragraph uses Pridi font.

Another paragraph still in Pridi.

| Table | Content |
|-------|---------|
| Also  | Pridi   |

![Image caption also uses Pridi](image.png)

<!-- {/font} -->

*เอกสารนี้สร้างโดย md2docx พร้อมฟอนต์ฝังตัว*
"#;

    let mut doc_config = DocumentConfig::default();
    doc_config.fonts = Some(md2docx::docx::FontConfig {
        default: Some("TH Sarabun New".to_string()),
        code: Some("Consolas".to_string()),
        normal_size: Some(28),  // 14pt
        normal_color: Some("000000".to_string()),
        h1_color: Some("000080".to_string()),
        caption_size: Some(24), // 12pt
        caption_color: Some("000000".to_string()),
        code_size: Some(20), // 10pt
    });
    // Just set embed_dir — fonts are automatically scanned, filtered, and embedded
    doc_config.embed_dir = Some(font_dir.to_path_buf());
    doc_config.toc.enabled = false;

    let docx_bytes = markdown_to_docx_with_templates(
        markdown,
        Language::Thai,
        &doc_config,
        None,
        &PlaceholderContext::default(),
    )?;

    let output_path = "output/font_embedding_test.docx";
    std::fs::create_dir_all("output")?;
    std::fs::write(output_path, &docx_bytes)?;

    println!("\n=== Output ===");
    println!("Written to: {}", output_path);
    println!("File size: {} bytes", std::fs::metadata(output_path)?.len());

    // Show what's inside the DOCX
    let reader = std::io::Cursor::new(std::fs::read(output_path)?);
    let mut archive = zip::ZipArchive::new(reader)?;
    println!("\n=== DOCX contents ===");
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        if file.name().contains("font") || file.name().contains("Font") {
            println!("  {} ({} bytes)", file.name(), file.size());
        }
    }

    Ok(())
}
