# Library API Reference {#ch08}

This chapter documents the Rust library API for embedding md2docx in your own applications. All public types, functions, and traits are covered with practical examples.

บทนี้เป็นเอกสารอ้างอิง API ไลบรารี Rust สำหรับการฝัง md2docx ในแอปพลิเคชันของคุณเอง ครอบคลุมประเภท ฟังก์ชัน และเทรตสาธารณะทั้งหมดพร้อมตัวอย่างที่ใช้ได้จริง

---

## Overview {#ch08-overview}

### English

The md2docx library provides a programmatic API for converting Markdown to DOCX. You can use it to:

- Build documentation generators
- Create custom CLI tools
- Integrate with web services
- Automate document workflows

### ภาษาไทย

ไลบรารี md2docx ให้ API แบบโปรแกรมสำหรับการแปลง Markdown เป็น DOCX คุณสามารถใช้เพื่อ:

- สร้างเครื่องมือสร้างเอกสาร
- สร้างเครื่องมือ CLI แบบกำหนดเอง
- ผสานรวมกับเว็บเซอร์วิส
- ทำงานเอกสารอัตโนมัติ

### Adding as Dependency / เพิ่มเป็น Dependency

Add to your `Cargo.toml`:

```toml
[dependencies]
md2docx = "0.1"
```

Or with specific features:

```toml
[dependencies]
md2docx = { version = "0.1", features = ["cli"] }
```

### Feature Flags / ฟีเจอร์แฟล็ก

| Feature | Description | Default |
|---------|-------------|---------|
| `cli` | Command-line interface support | Yes |
| `wasm` | WebAssembly bindings | No |
| `mermaid-cli` | Mermaid CLI rendering fallback | No |
| `thai-linebreak` | Thai word segmentation with ICU | No |

---

## Main Types {#ch08-main-types}

### Document / เอกสาร

The `Document` struct is the primary interface for document generation.

โครงสร้าง `Document` เป็นอินเตอร์เฟซหลักสำหรับการสร้างเอกสาร

```rust
use md2docx::Document;

// Simple conversion
let docx_bytes = Document::from_markdown("# Hello\n\nWorld")?;

// With builder pattern
let docx_bytes = Document::builder()
    .title("My Document")
    .author("John Doe")
    .add_markdown("# Chapter 1")
    .build()?;
```

### Config / การตั้งค่า

Configuration struct with nested sub-configs for fine-grained control.

โครงสร้างการตั้งค่าพร้อม sub-configs ที่ซ้อนกันสำหรับการควบคุมละเอียด

```rust
use md2docx::Config;

// Load from file
let config = Config::from_file("md2docx.toml")?;

// Or build programmatically
let config = Config::builder()
    .document(|d| d
        .title("My Manual")
        .author("Team")
        .language(Language::English))
    .toc(|t| t
        .enabled(true)
        .depth(3))
    .page_numbers(|p| p
        .enabled(true)
        .skip_cover(true))
    .build()?;
```

### Config Sub-structs / โครงสร้างย่อยของ Config

```rust
// DocumentConfig - metadata and structure
pub struct DocumentConfig {
    pub title: Option<String>,
    pub author: Option<String>,
    pub date: Option<String>,
    pub language: Language,
    pub toc: bool,
    pub page_numbers: bool,
}

// TocConfig - table of contents settings
pub struct TocConfig {
    pub enabled: bool,
    pub depth: u8,           // 1-6
    pub title: String,       // "Table of Contents" or "สารบัญ"
}

// PageNumberConfig - page numbering
pub struct PageNumberConfig {
    pub enabled: bool,
    pub skip_cover: bool,
    pub skip_chapter_first: bool,
    pub format: PageNumberFormat,
}

// TemplateConfig - template settings
pub struct TemplateConfig {
    pub file: Option<PathBuf>,
    pub validate: bool,
}

// CodeConfig - code block settings
pub struct CodeConfig {
    pub theme: CodeTheme,
    pub show_filename: bool,
    pub show_line_numbers: bool,
    pub highlight_lines: bool,
}

// ImageConfig - image handling
pub struct ImageConfig {
    pub max_width: String,   // "100%" or "600px"
    pub default_dpi: u32,
    pub figure_prefix: String,
    pub auto_caption: bool,
}
```

### Template / แม่แบบ

The `Template` struct represents a loaded DOCX template.

โครงสร้าง `Template` แทนแม่แบบ DOCX ที่โหลดแล้ว

```rust
use md2docx::Template;

// Load from file
let template = Template::from_file("custom-reference.docx")?;

// Or from bytes
let template_bytes = std::fs::read("template.docx")?;
let template = Template::from_bytes(&template_bytes)?;

// Validate
let report = template.validate()?;
if report.is_valid() {
    println!("Template is valid");
}
```

### Error / ข้อผิดพลาด

The `Error` enum covers all possible error conditions.

เอนัม `Error` ครอบคลุมเงื่อนไขข้อผิดพลาดที่เป็นไปได้ทั้งหมด

```rust
use md2docx::Error;

match result {
    Ok(docx) => { /* use docx */ }
    Err(Error::Parse(msg)) => eprintln!("Parse error: {}", msg),
    Err(Error::Io(e)) => eprintln!("IO error: {}", e),
    Err(Error::Template(msg)) => eprintln!("Template error: {}", msg),
    Err(Error::Config(msg)) => eprintln!("Config error: {}", msg),
    Err(e) => eprintln!("Other error: {}", e),
}
```

**Error Variants:**

| Variant | Description |
|---------|-------------|
| `Parse(String)` | Markdown parsing failed |
| `Io(std::io::Error)` | File/IO operation failed |
| `Template(String)` | Template error (missing styles, etc.) |
| `Config(String)` | Configuration error |
| `Validation(ValidationReport)` | Template validation failed |
| `Mermaid(String)` | Mermaid diagram rendering failed |
| `NotImplemented(String)` | Feature not yet implemented |

---

## Primary Functions {#ch08-primary-functions}

### from_markdown / จาก Markdown

Convert a Markdown string to DOCX bytes.

แปลงสตริง Markdown เป็นไบต์ DOCX

```rust
use md2docx::Document;

pub fn from_markdown(markdown: &str) -> Result<Vec<u8>, Error>
```

**Example / ตัวอย่าง:**

```rust
let markdown = r#"
# Hello World

This is a **bold** paragraph.

## Code Example

```rust
fn main() {
    println!("Hello!");
}
```
"#;

let docx_bytes = Document::from_markdown(markdown)?;
std::fs::write("output.docx", docx_bytes)?;
```

### from_file / จากไฟล์

Convert a Markdown file to DOCX bytes.

แปลงไฟล์ Markdown เป็นไบต์ DOCX

```rust
use std::path::Path;
use md2docx::Document;

pub fn from_file(path: &Path) -> Result<Vec<u8>, Error>
```

**Example / ตัวอย่าง:**

```rust
use std::path::Path;

let docx_bytes = Document::from_file(Path::new("README.md"))?;
std::fs::write("README.docx", docx_bytes)?;
```

### from_directory / จากไดเรกทอรี

Build a document from a project directory with configuration.

สร้างเอกสารจากไดเรกทอรีโครงการพร้อมการตั้งค่า

```rust
use std::path::Path;
use md2docx::{Document, Config};

pub fn from_directory(path: &Path, config: Config) -> Result<Vec<u8>, Error>
```

**Example / ตัวอย่าง:**

```rust
use md2docx::{Config, Language};

let config = Config::builder()
    .document(|d| d
        .title("Project Manual")
        .author("Dev Team")
        .language(Language::English))
    .toc(|t| t.enabled(true))
    .build()?;

let docx_bytes = Document::from_directory(
    Path::new("./docs/"),
    config
)?;

std::fs::write("manual.docx", docx_bytes)?;
```

### dump_template / สร้างแม่แบบ

Generate a default template DOCX file.

สร้างไฟล์แม่แบบ DOCX เริ่มต้น

```rust
use md2docx::{dump_template, Language};

pub fn dump_template(lang: Language) -> Result<Vec<u8>, Error>
```

**Example / ตัวอย่าง:**

```rust
// Generate English template
let template_bytes = dump_template(Language::English)?;
std::fs::write("template-en.docx", template_bytes)?;

// Generate Thai template
let template_bytes = dump_template(Language::Thai)?;
std::fs::write("template-th.docx", template_bytes)?;
```

### validate_template / ตรวจสอบแม่แบบ

Validate a template file and return a detailed report.

ตรวจสอบไฟล์แม่แบบและส่งคืนรายงานโดยละเอียด

```rust
use std::path::Path;
use md2docx::{validate_template, ValidationReport};

pub fn validate_template(path: &Path) -> Result<ValidationReport, Error>
```

**Example / ตัวอย่าง:**

```rust
let report = validate_template(Path::new("my-template.docx"))?;

println!("Required styles: {}/{}", 
    report.required_present(), 
    report.required_total());
println!("Recommended styles: {}/{}", 
    report.recommended_present(), 
    report.recommended_total());

if !report.is_valid() {
    for missing in report.missing_required() {
        eprintln!("Missing required style: {}", missing);
    }
}
```

---

## Builder Pattern {#ch08-builder-pattern}

### DocumentBuilder / ตัวสร้างเอกสาร

The builder pattern provides a fluent interface for document construction.

รูปแบบ builder ให้อินเตอร์เฟซ fluent สำหรับการสร้างเอกสาร

```rust
use md2docx::{Document, Template, Language};

let docx_bytes = Document::builder()
    // Metadata
    .title("API Documentation")
    .subtitle("Version 1.0")
    .author("Engineering Team")
    .date("2024-01-15")
    .language(Language::English)
    
    // Content
    .add_markdown("# Introduction\n\nWelcome to the API.")
    .add_markdown("## Authentication\n\nUse API keys.")
    .add_file("endpoints.md")?
    
    // Template
    .template(Template::from_file("template.docx")?)
    
    // Options
    .toc(true)
    .toc_depth(3)
    .page_numbers(true)
    
    // Build
    .build()?;
```

### Chaining Methods / เมธอดการเชื่อมโยง

| Method | Description |
|--------|-------------|
| `title(s: &str)` | Set document title |
| `subtitle(s: &str)` | Set document subtitle |
| `author(s: &str)` | Set document author |
| `date(s: &str)` | Set document date |
| `language(lang: Language)` | Set document language |
| `add_markdown(md: &str)` | Add markdown content |
| `add_file(path: &Path)` | Add markdown file |
| `template(t: Template)` | Set template |
| `toc(enabled: bool)` | Enable/disable TOC |
| `toc_depth(d: u8)` | Set TOC depth (1-6) |
| `page_numbers(enabled: bool)` | Enable/disable page numbers |

---

## Complete Examples {#ch08-complete-examples}

### Example 1: Simple Single File / ไฟล์เดี่ยวแบบง่าย

```rust
use md2docx::Document;
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Read markdown
    let markdown = fs::read_to_string("README.md")?;
    
    // Convert to DOCX
    let docx_bytes = Document::from_markdown(&markdown)?;
    
    // Write output
    fs::write("README.docx", docx_bytes)?;
    
    println!("Converted README.md to README.docx");
    Ok(())
}
```

### Example 2: Multi-File Project / โครงการหลายไฟล์

```rust
use md2docx::{Document, Config, Language};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Build configuration
    let config = Config::builder()
        .document(|d| d
            .title("Software Manual")
            .author("Documentation Team")
            .language(Language::English))
        .toc(|t| t
            .enabled(true)
            .depth(3)
            .title("Table of Contents".to_string()))
        .page_numbers(|p| p
            .enabled(true)
            .skip_cover(true))
        .build()?;
    
    // Build from directory
    let docx_bytes = Document::from_directory(
        Path::new("./docs/"),
        config
    )?;
    
    // Save output
    std::fs::write("manual.docx", docx_bytes)?;
    
    println!("Documentation built successfully!");
    Ok(())
}
```

### Example 3: With Custom Template / พร้อมแม่แบบกำหนดเอง

```rust
use md2docx::{Document, Template, Language};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load and validate template
    let template = Template::from_file(Path::new("company-template.docx"))?;
    
    let report = template.validate()?;
    if !report.is_valid() {
        eprintln!("Template validation failed!");
        for style in report.missing_required() {
            eprintln!("  Missing: {}", style);
        }
        return Err("Invalid template".into());
    }
    
    // Build document with template
    let docx_bytes = Document::builder()
        .title("Annual Report")
        .author("Finance Team")
        .language(Language::English)
        .template(template)
        .add_file(Path::new("report.md"))?
        .toc(true)
        .page_numbers(true)
        .build()?;
    
    std::fs::write("report.docx", docx_bytes)?;
    println!("Report generated with custom template!");
    
    Ok(())
}
```

### Example 4: Error Handling / การจัดการข้อผิดพลาด

```rust
use md2docx::{Document, Error, Template};
use std::path::Path;

fn convert_document(input: &Path, output: &Path) -> Result<(), Error> {
    // Attempt conversion with comprehensive error handling
    let docx_bytes = match Document::from_file(input) {
        Ok(bytes) => bytes,
        Err(Error::Io(e)) => {
            eprintln!("Failed to read input file: {}", e);
            return Err(Error::Io(e));
        }
        Err(Error::Parse(msg)) => {
            eprintln!("Failed to parse markdown: {}", msg);
            return Err(Error::Parse(msg));
        }
        Err(e) => {
            eprintln!("Unexpected error: {}", e);
            return Err(e);
        }
    };
    
    // Write output with error handling
    match std::fs::write(output, docx_bytes) {
        Ok(_) => {
            println!("Successfully created: {}", output.display());
            Ok(())
        }
        Err(e) => {
            eprintln!("Failed to write output: {}", e);
            Err(Error::Io(e))
        }
    }
}

fn main() {
    if let Err(e) = convert_document(
        Path::new("input.md"),
        Path::new("output.docx")
    ) {
        std::process::exit(1);
    }
}
```

---

## WASM Usage {#ch08-wasm-usage}

### English

> **Note**: WASM bindings are currently a stub and not fully implemented. This section documents the planned API.

WebAssembly support allows md2docx to run in browsers and serverless environments.

### ภาษาไทย

> **หมายเหตุ**: WASM bindings เป็นข้อกำหนดเท่านั้นและยังไม่ได้รับการพัฒนาเต็มรูปแบบ ส่วนนี้เป็นเอกสาร API ที่วางแผนไว้

การรองรับ WebAssembly ช่วยให้ md2docx ทำงานในเบราว์เซอร์และสภาพแวดล้อมแบบ serverless

### Planned API / API ที่วางแผนไว้

```javascript
// JavaScript usage (planned)
import init, { markdownToDocx } from 'md2docx';

await init();

// Simple conversion
const markdown = "# Hello\n\nWorld";
const docxBytes = markdownToDocx(markdown);

// Download
const blob = new Blob([docxBytes], { 
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
});
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'document.docx';
a.click();
```

### Feature Flag / ฟีเจอร์แฟล็ก

```toml
[dependencies]
md2docx = { version = "0.1", features = ["wasm"] }
```

### Current Status / สถานะปัจจุบัน

| Feature | Status |
|---------|--------|
| Core conversion | ⚠️ Stub only |
| Template support | ❌ Not implemented |
| Config support | ❌ Not implemented |
| File system access | ❌ Not available in WASM |

---

## Quick Reference / อ้างอิงด่วน

### Common Patterns / รูปแบบทั่วไป

```rust
// Quick single file
let bytes = Document::from_markdown("# Hello")?;

// With metadata
let bytes = Document::builder()
    .title("Title")
    .author("Author")
    .add_markdown(content)
    .build()?;

// From directory
let config = Config::from_file("config.toml")?;
let bytes = Document::from_directory(path, config)?;

// With template
let template = Template::from_file("template.docx")?;
let bytes = Document::builder()
    .template(template)
    .add_file("doc.md")?
    .build()?;
```

### Type Summary / สรุปประเภท

| Type | Purpose |
|------|---------|
| `Document` | Main conversion interface |
| `Config` | Configuration container |
| `Template` | Loaded DOCX template |
| `ValidationReport` | Template validation results |
| `Error` | Error enum |
| `Language` | Language enum (English, Thai) |

### Import Paths / พาธการนำเข้า

```rust
// Main types
use md2docx::{Document, Config, Template};

// Utility functions
use md2docx::{dump_template, validate_template};

// Enums
use md2docx::{Language, Error};
```
