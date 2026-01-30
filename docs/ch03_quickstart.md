---
title: "Quick Start / เริ่มต้นอย่างรวดเร็ว"
language: th
---

# Chapter 3: Quick Start / บทที่ 3: เริ่มต้นอย่างรวดเร็ว {#ch03}

Get started with md2docx in just 5 minutes!

เริ่มต้นใช้งาน md2docx ใน 5 นาที!

## Your First Document / เอกสารแรกของคุณ

### Step 1: Create a Markdown File / ขั้นตอนที่ 1: สร้างไฟล์ Markdown

Create a file named `hello.md`:

สร้างไฟล์ชื่อ `hello.md`:

```markdown
# Hello World / สวัสดีชาวโลก

This is my first document created with md2docx.

นี่คือเอกสารแรกของฉันที่สร้างด้วย md2docx

## Features / คุณสมบัติ

- Easy to use / ใช้งานง่าย
- Supports Thai / รองรับภาษาไทย
- Professional output / ผลลัพธ์มืออาชีพ

## Code Example / ตัวอย่างโค้ด

```rust
fn main() {
    println!("สวัสดีชาวโลก!");
}
```

> **Note / หมายเหตุ:** md2docx makes documentation easy!
> md2docx ทำให้การเขียนเอกสารเป็นเรื่องง่าย!
```

### Step 2: Convert to DOCX / ขั้นตอนที่ 2: แปลงเป็น DOCX

Run the md2docx command:

รันคำสั่ง md2docx:

```bash
md2docx build -i hello.md -o hello.docx
```

### Step 3: Open the Result / ขั้นตอนที่ 3: เปิดผลลัพธ์

Open `hello.docx` in Microsoft Word or any compatible application.

เปิด `hello.docx` ใน Microsoft Word หรือแอปพลิเคชันที่รองรับ

Congratulations! You've created your first DOCX with md2docx.

ยินดีด้วย! คุณได้สร้าง DOCX แรกด้วย md2docx

## Working with Multiple Files / การทำงานกับหลายไฟล์

For larger documents with multiple chapters:

สำหรับเอกสารขนาดใหญ่ที่มีหลายบท:

### Directory Structure / โครงสร้างโฟลเดอร์

```
my-docs/
├── md2docx.toml      # Configuration / การตั้งค่า
├── cover.md          # Cover page / หน้าปก
├── ch01_intro.md     # Chapter 1 / บทที่ 1
├── ch02_setup.md     # Chapter 2 / บทที่ 2
├── ch03_usage.md     # Chapter 3 / บทที่ 3
└── assets/           # Images / รูปภาพ
    └── logo.png
```

### Configuration File / ไฟล์การตั้งค่า

Create `md2docx.toml`:

สร้าง `md2docx.toml`:

```toml
[document]
title = "My Documentation"
author = "Your Name"
language = "th"

[toc]
enabled = true
depth = 2

[chapters]
pattern = "ch*_*.md"
sort = "numeric"
```

### Build the Document / Build เอกสาร

```bash
md2docx build -d ./my-docs/ -o output.docx
```

md2docx will automatically:

md2docx จะทำสิ่งเหล่านี้โดยอัตโนมัติ:

- Find and sort chapter files / ค้นหาและเรียงลำดับไฟล์บท
- Include the cover page first / ใส่หน้าปกก่อน
- Generate table of contents / สร้างสารบัญ
- Merge everything into one DOCX / รวมทุกอย่างเป็น DOCX เดียว

## Common Use Cases / กรณีใช้งานทั่วไป

### Technical Documentation / เอกสารทางเทคนิค

```bash
# Build with table of contents / Build พร้อมสารบัญ
md2docx build -d ./docs/ -o manual.docx --toc
```

### Reports / รายงาน

```bash
# Build with page numbers / Build พร้อมหมายเลขหน้า
md2docx build -d ./reports/ -o report.docx --page-numbers
```

### Books / หนังสือ

```bash
# Build with all features / Build พร้อมฟีเจอร์ทั้งหมด
md2docx build -d ./book/ -o book.docx --toc --page-numbers
```

## Quick Tips / เคล็ดลับ

### 1. Use Frontmatter / ใช้ Frontmatter

Add metadata to your Markdown files:

เพิ่ม metadata ในไฟล์ Markdown:

```markdown
---
title: "Chapter Title"
skip_toc: false
---

# Your Chapter Content
```

### 2. Include Images / ใส่รูปภาพ

```markdown
![Description / คำอธิบาย](assets/image.png)
```

### 3. Add Code with Syntax Highlighting / เพิ่มโค้ดพร้อม Syntax Highlighting

```markdown
```python,filename=example.py
def greet(name):
    print(f"Hello, {name}!")
```
```

### 4. Create Diagrams with Mermaid / สร้าง Diagram ด้วย Mermaid

```markdown
```mermaid
flowchart LR
    A[Start] --> B[Process] --> C[End]
```
```

## Next Steps / ขั้นตอนถัดไป

- **Chapter 4:** Learn all supported Markdown syntax
- **บทที่ 4:** เรียนรู้ไวยากรณ์ Markdown ที่รองรับทั้งหมด
- **Chapter 5:** Explore configuration options
- **บทที่ 5:** สำรวจตัวเลือกการตั้งค่า
