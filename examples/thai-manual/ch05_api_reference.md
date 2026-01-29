---
title: "การอ้างอิง API"
language: th
---

# บทที่ 5 การอ้างอิง API {#ch05}

บทนี้จะอธิบาย API ทั้งหมดของ md2docx รวมถึง Structs, Functions, และ Error Types ที่คุณสามารถใช้ในโปรแกรม Rust ของคุณ

## ภาพรวม API

md2docx มี API ที่ยืดหยุ่นและใช้งานง่าย คุณสามารถใช้ md2docx ได้ 2 วิธีหลัก:

1. **ใช้เป็น Library** - ฝังในโปรแกรม Rust ของคุณ
2. **ใช้เป็น CLI** - ใช้คำสั่ง command line

บทนี้จะเน้นที่การใช้เป็น Library

## Structs หลัก

### Document

`Document` เป็น Struct หลักที่ใช้สร้างเอกสาร DOCX

```rust
pub struct Document {
    title: String,
    config: Config,
    template: Option<Template>,
    chapters: Vec<Chapter>,
    // ... fields อื่นๆ
}
```

#### การสร้าง Document ใหม่

**วิธีที่ 1: จาก Markdown เดียว**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let markdown = "# สวัสดีชาวโลก\n\nนี่คือเอกสารแรกของฉัน";
    let docx = Document::from_markdown(markdown)?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**วิธีที่ 2: จากไฟล์**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let docx = Document::from_file("README.md")?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**วิธีที่ 3: จากโฟลเดอร์**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = Config::default();
    let docx = Document::from_directory("./docs/", config)?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**วิธีที่ 4: ใช้ Builder Pattern**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let docx = Document::builder()
        .title("คู่มือการใช้งาน")
        .author("ทีมพัฒนา")
        .language("th")
        .add_file("cover.md")?
        .add_file("ch01_introduction.md")?
        .add_file("ch02_installation.md")?
        .build()?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

#### Methods ของ Document

| Method | คำอธิบาย | ค่าที่คืน |
|--------|-----------|-----------|
| `from_markdown(markdown: &str)` | สร้าง Document จาก Markdown string | `Result<Document, Error>` |
| `from_file(path: &Path)` | สร้าง Document จากไฟล์ Markdown | `Result<Document, Error>` |
| `from_directory(path: &Path, config: Config)` | สร้าง Document จากโฟลเดอร์ | `Result<Document, Error>` |
| `builder()` | สร้าง DocumentBuilder | `DocumentBuilder` |
| `write_to(path: &Path)` | เขียนเอกสารเป็นไฟล์ DOCX | `Result<(), Error>` |
| `to_bytes()` | แปลงเป็น bytes | `Result<Vec<u8>, Error>` |

### Config

`Config` ใช้กำหนดการตั้งค่าทั้งหมดของเอกสาร

```rust
pub struct Config {
    pub document: DocumentConfig,
    pub template: TemplateConfig,
    pub output: OutputConfig,
    pub toc: TocConfig,
    pub page_numbers: PageNumbersConfig,
    pub header: HeaderConfig,
    pub footer: FooterConfig,
    pub fonts: FontsConfig,
    pub code: CodeConfig,
    pub images: ImagesConfig,
}
```

#### การสร้าง Config

**ใช้ค่าเริ่มต้น:**

```rust
use md2docx::Config;

fn main() {
    let config = Config::default();
}
```

**จากไฟล์ TOML:**

```rust
use md2docx::Config;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = Config::from_file("md2docx.toml")?;
    Ok(())
}
```

**สร้างเอง:**

```rust
use md2docx::Config;

fn main() {
    let config = Config {
        document: DocumentConfig {
            title: "คู่มือการใช้งาน".to_string(),
            author: "ทีมพัฒนา".to_string(),
            language: "th".to_string(),
            ..Default::default()
        },
        ..Default::default()
    };
}
```

#### Sub-configs

**DocumentConfig:**

```rust
pub struct DocumentConfig {
    pub title: String,
    pub subtitle: Option<String>,
    pub author: Option<String>,
    pub date: String,
    pub language: String,
}
```

**TemplateConfig:**

```rust
pub struct TemplateConfig {
    pub file: Option<String>,
    pub validate: bool,
}
```

**OutputConfig:**

```rust
pub struct OutputConfig {
    pub file: String,
    pub format: String,
}
```

**TocConfig:**

```rust
pub struct TocConfig {
    pub enabled: bool,
    pub depth: u8,
    pub title: Option<String>,
}
```

**PageNumbersConfig:**

```rust
pub struct PageNumbersConfig {
    pub enabled: bool,
    pub skip_cover: bool,
    pub skip_chapter_first: bool,
    pub format: String,
}
```

**HeaderConfig:**

```rust
pub struct HeaderConfig {
    pub left: String,
    pub center: String,
    pub right: String,
    pub skip_cover: bool,
}
```

**FooterConfig:**

```rust
pub struct FooterConfig {
    pub left: String,
    pub center: String,
    pub right: String,
    pub skip_cover: bool,
}
```

**FontsConfig:**

```rust
pub struct FontsConfig {
    pub default: String,
    pub thai: String,
    pub code: String,
    pub fallback: Vec<String>,
}
```

**CodeConfig:**

```rust
pub struct CodeConfig {
    pub theme: String,
    pub show_filename: bool,
    pub show_line_numbers: bool,
    pub font: String,
}
```

**ImagesConfig:**

```rust
pub struct ImagesConfig {
    pub max_width: String,
    pub default_dpi: u32,
    pub figure_prefix: String,
    pub auto_caption: bool,
}
```

### Template

`Template` ใช้โหลดและใช้ Template จากไฟล์ DOCX

```rust
pub struct Template {
    styles: HashMap<String, Style>,
    // ... fields อื่นๆ
}
```

#### การสร้าง Template

**จากไฟล์:**

```rust
use md2docx::Template;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let template = Template::from_file("custom-reference.docx")?;
    Ok(())
}
```

**จาก bytes:**

```rust
use md2docx::Template;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read("custom-reference.docx")?;
    let template = Template::from_bytes(&bytes)?;
    Ok(())
}
```

#### Methods ของ Template

| Method | คำอธิบาย | ค่าที่คืน |
|--------|-----------|-----------|
| `from_file(path: &Path)` | โหลด Template จากไฟล์ | `Result<Template, Error>` |
| `from_bytes(bytes: &[u8])` | โหลด Template จาก bytes | `Result<Template, Error>` |
| `get_style(id: &str)` | ดึง Style จาก ID | `Option<&Style>` |
| `validate()` | ตรวจสอบ Template | `Result<(), Error>` |

## Public Functions

### from_markdown

แปลง Markdown string เป็น DOCX

```rust
pub fn from_markdown(markdown: &str) -> Result<Vec<u8>, Error>
```

**ตัวอย่าง:**

```rust
use md2docx::from_markdown;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let markdown = "# สวัสดี\n\nนี่คือเอกสาร";
    let docx_bytes = from_markdown(markdown)?;
    std::fs::write("output.docx", docx_bytes)?;
    Ok(())
}
```

### from_file

แปลงไฟล์ Markdown เป็น DOCX

```rust
pub fn from_file(path: &Path) -> Result<Vec<u8>, Error>
```

**ตัวอย่าง:**

```rust
use md2docx::from_file;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let docx_bytes = from_file(Path::new("README.md"))?;
    std::fs::write("output.docx", docx_bytes)?;
    Ok(())
}
```

### from_directory

แปลงโฟลเดอร์ Markdown เป็น DOCX

```rust
pub fn from_directory(path: &Path, config: Config) -> Result<Vec<u8>, Error>
```

**ตัวอย่าง:**

```rust
use md2docx::{from_directory, Config};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = Config::default();
    let docx_bytes = from_directory(Path::new("./docs/"), config)?;
    std::fs::write("output.docx", docx_bytes)?;
    Ok(())
}
```

### dump_template

สร้าง Template เริ่มต้น

```rust
pub fn dump_template(lang: Language) -> Result<Vec<u8>, Error>
```

**ตัวอย่าง:**

```rust
use md2docx::{dump_template, Language};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let template_bytes = dump_template(Language::Thai)?;
    std::fs::write("template.docx", template_bytes)?;
    Ok(())
}
```

### validate_template

ตรวจสอบ Template

```rust
pub fn validate_template(path: &Path) -> Result<ValidationReport, Error>
```

**ตัวอย่าง:**

```rust
use md2docx::validate_template;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let report = validate_template(Path::new("template.docx"))?;
    println!("Template is valid: {}", report.is_valid());
    Ok(())
}
```

## Builder Pattern

### DocumentBuilder

ใช้สร้าง Document แบบ step-by-step

```rust
pub struct DocumentBuilder {
    title: Option<String>,
    author: Option<String>,
    language: String,
    config: Config,
    template: Option<Template>,
    chapters: Vec<Chapter>,
}
```

#### Methods ของ DocumentBuilder

| Method | คำอธิบาย | ค่าที่คืน |
|--------|-----------|-----------|
| `title(title: &str)` | กำหนดชื่อเอกสาร | `&mut Self` |
| `author(author: &str)` | กำหนดผู้แต่ง | `&mut Self` |
| `language(lang: &str)` | กำหนดภาษา | `&mut Self` |
| `config(config: Config)` | กำหนด Config | `&mut Self` |
| `template(template: Template)` | กำหนด Template | `&mut Self` |
| `add_file(path: &Path)` | เพิ่มไฟล์ Markdown | `Result<&mut Self, Error>` |
| `add_markdown(name: &str, content: &str)` | เพิ่ม Markdown string | `&mut Self` |
| `build()` | สร้าง Document | `Result<Document, Error>` |

#### ตัวอย่างการใช้งาน

**ตัวอย่างที่ 1: พื้นฐาน**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let docx = Document::builder()
        .title("คู่มือการใช้งาน")
        .author("ทีมพัฒนา")
        .language("th")
        .build()?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**ตัวอย่างที่ 2: พร้อม Config**

```rust
use md2docx::{Document, Config};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = Config {
        document: md2docx::DocumentConfig {
            title: "คู่มือการใช้งาน".to_string(),
            ..Default::default()
        },
        ..Default::default()
    };

    let docx = Document::builder()
        .config(config)
        .build()?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**ตัวอย่างที่ 3: พร้อม Template**

```rust
use md2docx::{Document, Template};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let template = Template::from_file("custom-reference.docx")?;

    let docx = Document::builder()
        .template(template)
        .add_file("cover.md")?
        .add_file("ch01_introduction.md")?
        .build()?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**ตัวอย่างที่ 4: พร้อม Markdown string**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let docx = Document::builder()
        .title("เอกสารแบบกำหนดเอง")
        .add_markdown("cover", "# หน้าปก\n\nเนื้อหาหน้าปก")
        .add_markdown("ch01", "# บทที่ 1\n\nเนื้อหาบทที่ 1")
        .build()?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

## Error Types

### Error Enum

md2docx ใช้ `thiserror` สำหรับการจัดการ Error

```rust
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("Failed to parse markdown: {0}")]
    Parse(String),

    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Template error: {0}")]
    Template(String),

    #[error("Config error: {0}")]
    Config(String),

    #[error("Image error: {0}")]
    Image(String),

    #[error("Mermaid error: {0}")]
    Mermaid(String),
}
```

### การจัดการ Error

**ตัวอย่างที่ 1: พื้นฐาน**

```rust
use md2docx::Document;

fn main() {
    match Document::from_file("README.md") {
        Ok(docx) => {
            if let Err(e) = docx.write_to("output.docx") {
                eprintln!("Error writing file: {}", e);
            }
        }
        Err(e) => {
            eprintln!("Error reading file: {}", e);
        }
    }
}
```

**ตัวอย่างที่ 2: ใช้ `?` operator**

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let docx = Document::from_file("README.md")?;
    docx.write_to("output.docx")?;
    Ok(())
}
```

**ตัวอย่างที่ 3: จัดการ Error เฉพาะ**

```rust
use md2docx::{Document, Error};

fn main() {
    match Document::from_file("README.md") {
        Ok(docx) => {
            if let Err(e) = docx.write_to("output.docx") {
                match e {
                    Error::Io(io_err) => {
                        eprintln!("IO error: {}", io_err);
                    }
                    Error::Parse(parse_err) => {
                        eprintln!("Parse error: {}", parse_err);
                    }
                    _ => {
                        eprintln!("Other error: {}", e);
                    }
                }
            }
        }
        Err(e) => {
            eprintln!("Error: {}", e);
        }
    }
}
```

## ตัวอย่างการใช้งานแบบเต็ม

### ตัวอย่างที่ 1: แปลงไฟล์เดียว

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // อ่านไฟล์ Markdown
    let docx = Document::from_file("README.md")?;

    // เขียนเป็น DOCX
    docx.write_to("output.docx")?;

    println!("สร้างไฟล์ output.docx เรียบร้อยแล้ว");
    Ok(())
}
```

### ตัวอย่างที่ 2: แปลงโฟลเดอร์พร้อม Config

```rust
use md2docx::{Document, Config};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // สร้าง Config
    let config = Config {
        document: md2docx::DocumentConfig {
            title: "คู่มือการใช้งาน".to_string(),
            author: Some("ทีมพัฒนา".to_string()),
            language: "th".to_string(),
            ..Default::default()
        },
        template: md2docx::TemplateConfig {
            file: Some("custom-reference.docx".to_string()),
            validate: true,
        },
        toc: md2docx::TocConfig {
            enabled: true,
            depth: 3,
            title: Some("สารบัญ".to_string()),
        },
        page_numbers: md2docx::PageNumbersConfig {
            enabled: true,
            skip_cover: true,
            skip_chapter_first: true,
            format: "หน้า {n}".to_string(),
        },
        ..Default::default()
    };

    // แปลงโฟลเดอร์
    let docx = Document::from_directory("./docs/", config)?;

    // เขียนเป็น DOCX
    docx.write_to("output/manual.docx")?;

    println!("สร้างไฟล์ output/manual.docx เรียบร้อยแล้ว");
    Ok(())
}
```

### ตัวอย่างที่ 3: ใช้ Builder Pattern

```rust
use md2docx::{Document, Template};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // โหลด Template
    let template = Template::from_file("custom-reference.docx")?;

    // สร้าง Document ด้วย Builder
    let docx = Document::builder()
        .title("คู่มือการใช้งานระบบ")
        .author("ทีมพัฒนา")
        .language("th")
        .template(template)
        .add_file("cover.md")?
        .add_file("ch01_introduction.md")?
        .add_file("ch02_installation.md")?
        .add_file("ch03_configuration.md")?
        .build()?;

    // เขียนเป็น DOCX
    docx.write_to("output/manual.docx")?;

    println!("สร้างไฟล์ output/manual.docx เรียบร้อยแล้ว");
    Ok(())
}
```

### ตัวอย่างที่ 4: สร้างเอกสารจาก Markdown string

```rust
use md2docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // สร้างเนื้อหา
    let cover = r#"
# คู่มือการใช้งาน

## เวอร์ชัน 2.0

**ทีมพัฒนา**

มกราคม 2568
"#;

    let ch01 = r#"
# บทที่ 1 บทนำ

ยินดีต้อนรับสู่คู่มือการใช้งานระบบ

## การเริ่มต้น

เริ่มต้นใช้งานระบบโดยทำตามขั้นตอนต่อไปนี้:

1. ลงทะเบียนบัญชี
2. ยืนยันอีเมล
3. เริ่มใช้งาน
"#;

    let ch02 = r#"
# บทที่ 2 การติดตั้ง

## ความต้องการของระบบ

- ระบบปฏิบัติการ: Windows, macOS, Linux
- หน่วยความจำ: 512MB ขึ้นไป
- พื้นที่จัดเก็บ: 50MB

## ขั้นตอนการติดตั้ง

ดูรายละเอียดในเอกสารการติดตั้ง
"#;

    // สร้าง Document
    let docx = Document::builder()
        .title("คู่มือการใช้งานระบบ")
        .add_markdown("cover", cover)
        .add_markdown("ch01", ch01)
        .add_markdown("ch02", ch02)
        .build()?;

    // เขียนเป็น DOCX
    docx.write_to("output/manual.docx")?;

    println!("สร้างไฟล์ output/manual.docx เรียบร้อยแล้ว");
    Ok(())
}
```

### ตัวอย่างที่ 5: จัดการ Error อย่างละเอียด

```rust
use md2docx::{Document, Error};

fn main() {
    match create_document() {
        Ok(_) => println!("สร้างเอกสารสำเร็จ"),
        Err(e) => handle_error(e),
    }
}

fn create_document() -> Result<(), Error> {
    let docx = Document::from_file("README.md")?;
    docx.write_to("output.docx")?;
    Ok(())
}

fn handle_error(error: Error) {
    match error {
        Error::Io(io_err) => {
            eprintln!("ข้อผิดพลาด IO: {}", io_err);
            eprintln!("โปรดตรวจสอบว่าไฟล์มีอยู่จริงและมีสิทธิ์เขียน");
        }
        Error::Parse(parse_err) => {
            eprintln!("ข้อผิดพลาดการแปลง Markdown: {}", parse_err);
            eprintln!("โปรดตรวจสอบรูปแบบ Markdown");
        }
        Error::Template(template_err) => {
            eprintln!("ข้อผิดพลาด Template: {}", template_err);
            eprintln!("โปรดตรวจสอบไฟล์ Template");
        }
        Error::Config(config_err) => {
            eprintln!("ข้อผิดพลาด Config: {}", config_err);
            eprintln!("โปรดตรวจสอบไฟล์ md2docx.toml");
        }
        Error::Image(image_err) => {
            eprintln!("ข้อผิดพลาดรูปภาพ: {}", image_err);
            eprintln!("โปรดตรวจสอบไฟล์รูปภาพ");
        }
        Error::Mermaid(mermaid_err) => {
            eprintln!("ข้อผิดพลาด Mermaid: {}", mermaid_err);
            eprintln!("โปรดตรวจสอบโค้ด Mermaid");
        }
    }
}
```

## สรุปบทนี้

ในบทนี้คุณได้เรียนรู้ API ทั้งหมดของ md2docx:

- **Structs หลัก** - Document, Config, Template
- **Public Functions** - from_markdown, from_file, from_directory, dump_template, validate_template
- **Builder Pattern** - DocumentBuilder และการใช้งาน
- **Error Types** - Error Enum และการจัดการ
- **ตัวอย่างการใช้งาน** - 5 ตัวอย่างแบบเต็ม

ในบทถัดไป เราจะเรียนรู้เกี่ยวกับ **การแก้ไขปัญหาและ FAQ** เพื่อช่วยให้คุณแก้ไขปัญหาที่พบบ่อย