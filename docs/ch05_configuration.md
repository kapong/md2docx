# Configuration Reference {#ch05}

This chapter documents all configuration options for md2docx. The configuration file (`md2docx.toml`) controls how your documents are generated, from styling to structure.

บทนี้อธิบายตัวเลือกการตั้งค่าทั้งหมดสำหรับ md2docx ไฟล์การตั้งค่า (`md2docx.toml`) ควบคุมวิธีการสร้างเอกสารของคุณ ตั้งแต่การจัดรูปแบบไปจนถึงโครงสร้าง

---

## Overview {#ch05-overview}

### Configuration File Locations / ตำแหน่งไฟล์การตั้งค่า

md2docx searches for configuration files in the following order (first found wins):

md2docx จะค้นหาไฟล์การตั้งค่าตามลำดับดังนี้ (ไฟล์แรกที่พบจะถูกใช้):

1. Path specified via `--config` option
   - ตำแหน่งที่ระบุผ่านตัวเลือก `--config`
2. `md2docx.toml` in current directory
   - `md2docx.toml` ในไดเรกทอรีปัจจุบัน
3. `.md2docx.toml` in current directory (hidden file)
   - `.md2docx.toml` ในไดเรกทอรีปัจจุบัน (ไฟล์ซ่อน)
4. `md2docx.toml` in parent directories (up to project root)
   - `md2docx.toml` ในไดเรกทอรีหลัก (ขึ้นไปจนถึงรากโครงการ)

### Configuration Priority / ลำดับความสำคัญของการตั้งค่า

Settings are applied in this priority (highest to lowest):

การตั้งค่าจะถูกนำไปใช้ตามลำดับความสำคัญดังนี้ (จากสูงไปต่ำ):

1. **Command-line options** (e.g., `--title`, `--toc-depth`)
   - **ตัวเลือกบนบรรทัดคำสั่ง** (เช่น `--title`, `--toc-depth`)
2. **Environment variables** (e.g., `MD2DOCX_TITLE`)
   - **ตัวแปรสภาพแวดล้อม** (เช่น `MD2DOCX_TITLE`)
3. **Config file values**
   - **ค่าในไฟล์การตั้งค่า**
4. **Built-in defaults**
   - **ค่าเริ่มต้นภายใน**

---

## [document] Section {#ch05-document}

Document metadata and global settings.

ข้อมูลเมตาของเอกสารและการตั้งค่าทั่วไป

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `title` | string | `""` | Document title / ชื่อเอกสาร |
| `subtitle` | string | `""` | Document subtitle / คำบรรยายใต้ชื่อ |
| `author` | string | `""` | Author name / ชื่อผู้เขียน |
| `date` | string | `"auto"` | Date format or "auto" / รูปแบบวันที่หรือ "auto" |
| `language` | string | `"en"` | Document language (`en` or `th`) / ภาษาของเอกสาร |
| `version` | string | `""` | Document version / เวอร์ชันของเอกสาร |

### Examples / ตัวอย่าง

```toml
[document]
title = "My Software Manual"
subtitle = "Complete User Guide"
author = "John Doe"
date = "auto"  # Uses current date / ใช้วันที่ปัจจุบัน
language = "en"
version = "1.0.0"
```

```toml
[document]
title = "คู่มือซอฟต์แวร์"
subtitle = "คู่มือการใช้งานฉบับสมบูรณ์"
author = "สมชาย ใจดี"
date = "2024-01-15"  # Specific date / วันที่เฉพาะเจาะจง
language = "th"
version = "2.1.0"
```

### Date Formats / รูปแบบวันที่

- `"auto"` - Current date in localized format
  - วันที่ปัจจุบันในรูปแบบท้องถิ่น
- `"YYYY-MM-DD"` - ISO format (e.g., "2024-01-15")
  - รูปแบบ ISO
- Custom format with placeholders:
  - รูปแบบกำหนดเองด้วยตัวยึดตำแหน่ง:
  - `{year}`, `{month}`, `{day}` - Numbers
    - ตัวเลข
  - `{month_name}` - Full month name
    - ชื่อเดือนเต็ม
  - `{month_short}` - Abbreviated month name
    - ชื่อเดือนย่อ

---

## [template] Section {#ch05-template}

Template file configuration.

การตั้งค่าไฟล์แม่แบบ

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `file` | string | `""` | Path to template DOCX file / พาธไปยังไฟล์แม่แบบ DOCX |
| `dir` | string | `""` | Path to template directory / พาธไปยังไดเรกทอรีแม่แบบ |
| `validate` | boolean | `true` | Validate template on load / ตรวจสอบแม่แบบเมื่อโหลด |

### Template Directory Structure / โครงสร้างไดเรกทอรีแม่แบบ

When using `dir`, md2docx looks for these files:

เมื่อใช้ `dir` md2docx จะค้นหาไฟล์เหล่านี้:

```
template-dir/
├── styles.docx        # Style definitions / คำนิยาม Styles
├── cover.docx         # Cover page template / แม่แบบหน้าปก
├── header-footer.docx # Header/footer template / แม่แบบส่วนหัว/ท้าย
├── image.docx         # Image styling template / แม่แบบการจัดรูปแบบรูปภาพ
└── table.docx         # Table styling template / แม่แบบการจัดรูปแบบตาราง
```

### Examples / ตัวอย่าง

```toml
[template]
file = "custom-reference.docx"
validate = true
```

```toml
[template]
dir = "./templates/company-template/"
validate = true
```

---

## [output] Section {#ch05-output}

Output file configuration.

การตั้งค่าไฟล์เอาต์พุต

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `file` | string | `"output.docx"` | Output filename / ชื่อไฟล์เอาต์พุต |
| `format` | string | `"docx"` | Output format (currently only "docx") / รูปแบบเอาต์พุต |
| `timestamp` | boolean | `false` | Append timestamp to filename / เพิ่มประทับเวลาในชื่อไฟล์ |

### Filename Placeholders / ตัวยึดตำแหน่งในชื่อไฟล์

The `file` option supports these placeholders:

ตัวเลือก `file` รองรับตัวยึดตำแหน่งเหล่านี้:

| Placeholder | Description |
|-------------|-------------|
| `{date}` | Current date (YYYY-MM-DD) / วันที่ปัจจุบัน |
| `{time}` | Current time (HH-MM-SS) / เวลาปัจจุบัน |
| `{datetime}` | Date and time / วันที่และเวลา |
| `{version}` | Document version from [document] / เวอร์ชันเอกสาร |

### Examples / ตัวอย่าง

```toml
[output]
file = "manual-{date}.docx"  # manual-2024-01-15.docx
```

```toml
[output]
file = "{version}-documentation.docx"  # 1.0.0-documentation.docx
```

```toml
[output]
file = "output/{datetime}-build.docx"  # output/2024-01-15-14-30-00-build.docx
```

---

## [toc] Section {#ch05-toc}

Table of Contents configuration.

การตั้งค่าสารบัญ

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `enabled` | boolean | `true` | Include TOC / รวมสารบัญ |
| `depth` | integer | `3` | Maximum heading depth (1-6) / ความลึกสูงสุดของหัวข้อ |
| `title` | string | `"Table of Contents"` | TOC title / ชื่อสารบัญ |

### Examples / ตัวอย่าง

```toml
[toc]
enabled = true
depth = 3
title = "Table of Contents"
```

```toml
[toc]
enabled = true
depth = 2  # Only H1 and H2 / เฉพาะ H1 และ H2
title = "สารบัญ"  # Thai title / ชื่อภาษาไทย
```

---

## [fonts] Section {#ch05-fonts}

Font configuration for different scripts and contexts.

การตั้งค่าฟอนต์สำหรับสคริปต์และบริบทต่างๆ

### Options / ตัวเลือก

| Option | Type | Default (EN) | Default (TH) | Description |
|--------|------|--------------|--------------|-------------|
| `default` | string | `"Calibri"` | `"TH Sarabun New"` | Default font / ฟอนต์เริ่มต้น |
| `thai` | string | `"TH Sarabun New"` | `"TH Sarabun New"` | Thai script font / ฟอนต์ภาษาไทย |
| `code` | string | `"Consolas"` | `"Consolas"` | Monospace font / ฟอนต์แบบไม่มีเว้นวรรค |
| `fallback` | array | `["Arial Unicode MS"]` | `["Tahoma"]` | Fallback fonts / ฟอนต์สำรอง |

### Style-Based Font Options / ตัวเลือกฟอนต์ตามสไตล์

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `normal_based_size` | integer | `22` | Normal style size in half-points (11pt) / ขนาดสไตล์ปกติ |
| `h1_based_size` | integer | `32` | Heading 1 size (16pt) / ขนาดหัวข้อ 1 |
| `h2_based_size` | integer | `26` | Heading 2 size (13pt) / ขนาดหัวข้อ 2 |
| `h3_based_size` | integer | `24` | Heading 3 size (12pt) / ขนาดหัวข้อ 3 |
| `h1_based_color` | string | `"2E74B5"` | Heading 1 color (hex) / สีหัวข้อ 1 |
| `h2_based_color` | string | `"2E74B5"` | Heading 2 color (hex) / สีหัวข้อ 2 |

### Examples / ตัวอย่าง

```toml
[fonts]
default = "Arial"
thai = "TH Sarabun New"
code = "Courier New"
fallback = ["Arial Unicode MS", "Segoe UI"]
```

```toml
[fonts]
default = "TH Sarabun New"
thai = "TH Sarabun New"
code = "Consolas"
normal_based_size = 28  # 14pt for Thai readability / 14pt สำหรับการอ่านภาษาไทย
h1_based_size = 40      # 20pt / 20pt
h1_based_color = "C00000"  # Red headings / หัวข้อสีแดง
```

---

## [code] Section {#ch05-code}

Code block display configuration.

การตั้งค่าการแสดงบล็อกโค้ด

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `theme` | string | `"light"` | Color theme (`light`, `dark`) / ธีมสี |
| `show_filename` | boolean | `true` | Show filename header / แสดงส่วนหัวชื่อไฟล์ |
| `show_line_numbers` | boolean | `false` | Show line numbers / แสดงหมายเลขบรรทัด |
| `highlight_lines` | boolean | `true` | Enable line highlighting / เปิดใช้งานการไฮไลต์บรรทัด |
| `font` | string | `"Consolas"` | Code font / ฟอนต์โค้ด |
| `source_root` | string | `"."` | Base path for `{!code:...}` / พาธฐานสำหรับ `{!code:...}` |

### Examples / ตัวอย่าง

```toml
[code]
theme = "light"
show_filename = true
show_line_numbers = true
highlight_lines = true
font = "Fira Code"
source_root = "../src"
```

---

## [chapters] Section {#ch05-chapters}

Chapter file discovery configuration.

การตั้งค่าการค้นหาไฟล์บท

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `pattern` | string | `"ch*_*.md"` | Glob pattern for chapter files / รูปแบบ glob สำหรับไฟล์บท |
| `sort` | string | `"numeric"` | Sort method (`numeric`, `alphabetic`, `none`) / วิธีการเรียงลำดับ |
| `skip_cover` | boolean | `true` | Skip cover.md in chapter list / ข้าม cover.md ในรายการบท |

### Sort Methods / วิธีการเรียงลำดับ

- `"numeric"` - Sort by numeric prefix (ch01, ch02, ch10)
  - เรียงตามคำนำหน้าตัวเลข (ch01, ch02, ch10)
- `"alphabetic"` - Sort alphabetically by filename
  - เรียงตามตัวอักษรตามชื่อไฟล์
- `"none"` - Use filesystem order
  - ใช้ลำดับของระบบไฟล์

### Examples / ตัวอย่าง

```toml
[chapters]
pattern = "ch*_*.md"
sort = "numeric"
```

```toml
[chapters]
pattern = "section-*.md"  # Custom pattern / รูปแบบกำหนดเอง
sort = "alphabetic"
```

---

## [appendices] Section {#ch05-appendices}

Appendix file discovery configuration.

การตั้งค่าการค้นหาไฟล์ภาคผนวก

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `pattern` | string | `"ap*_*.md"` | Glob pattern for appendix files / รูปแบบ glob สำหรับไฟล์ภาคผนวก |
| `prefix` | string | `"Appendix"` | Prefix for appendix labels / คำนำหน้าสำหรับฉลากภาคผนวก |
| `sort` | string | `"alphabetic"` | Sort method / วิธีการเรียงลำดับ |

### Examples / ตัวอย่าง

```toml
[appendices]
pattern = "ap*_*.md"
prefix = "Appendix"
```

```toml
[appendices]
pattern = "appendix-*.md"
prefix = "ภาคผนวก"  # Thai prefix / คำนำหน้าภาษาไทย
sort = "numeric"
```

---

## [page_numbers] Section {#ch05-page-numbers}

Page numbering configuration.

การตั้งค่าหมายเลขหน้า

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `enabled` | boolean | `true` | Enable page numbers / เปิดใช้งานหมายเลขหน้า |
| `skip_cover` | boolean | `true` | No page number on cover / ไม่มีหมายเลขหน้าบนหน้าปก |
| `skip_chapter_first` | boolean | `true` | No page number on chapter first pages / ไม่มีหมายเลขหน้าบนหน้าแรกของบท |
| `format` | string | `"{n}"` | Page number format / รูปแบบหมายเลขหน้า |
| `start_at` | integer | `1` | Starting page number / หมายเลขหน้าเริ่มต้น |

### Format Placeholders / ตัวยึดตำแหน่งรูปแบบ

| Placeholder | Description |
|-------------|-------------|
| `{n}` | Current page number / หมายเลขหน้าปัจจุบัน |
| `{total}` | Total pages / จำนวนหน้าทั้งหมด |

### Examples / ตัวอย่าง

```toml
[page_numbers]
enabled = true
skip_cover = true
skip_chapter_first = true
format = "Page {n} of {total}"
```

```toml
[page_numbers]
enabled = true
format = "{n}"  # Simple numbering / การนับแบบง่าย
start_at = 1
```

```toml
[page_numbers]
enabled = true
format = "หน้า {n} จาก {total}"  # Thai format / รูปแบบภาษาไทย
```

---

## [header] Section {#ch05-header}

Header configuration.

การตั้งค่าส่วนหัว

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `left` | string | `"{title}"` | Left header content / เนื้อหาส่วนหัวซ้าย |
| `center` | string | `""` | Center header content / เนื้อหาส่วนหัวกลาง |
| `right` | string | `"{chapter}"` | Right header content / เนื้อหาส่วนหัวขวา |
| `skip_cover` | boolean | `true` | No header on cover page / ไม่มีส่วนหัวบนหน้าปก |

### Header Placeholders / ตัวยึดตำแหน่งส่วนหัว

| Placeholder | Description |
|-------------|-------------|
| `{title}` | Document title / ชื่อเอกสาร |
| `{chapter}` | Current chapter name / ชื่อบทปัจจุบัน |
| `{author}` | Author name / ชื่อผู้เขียน |
| `{date}` | Current date / วันที่ปัจจุบัน |
| `{version}` | Document version / เวอร์ชันเอกสาร |

### Examples / ตัวอย่าง

```toml
[header]
left = "{title}"
center = ""
right = "{chapter}"
skip_cover = true
```

```toml
[header]
left = "My Company"
center = "{title}"
right = "Confidential"
```

---

## [footer] Section {#ch05-footer}

Footer configuration.

การตั้งค่าส่วนท้าย

### Options / ตัวเลือก

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `left` | string | `""` | Left footer content / เนื้อหาส่วนท้ายซ้าย |
| `center` | string | `"{page}"` | Center footer content / เนื้อหาส่วนท้ายกลาง |
| `right` | string | `""` | Right footer content / เนื้อหาส่วนท้ายขวา |
| `skip_cover` | boolean | `true` | No footer on cover page / ไม่มีส่วนท้ายบนหน้าปก |

### Footer Placeholders / ตัวยึดตำแหน่งส่วนท้าย

Same as header placeholders, plus:

เหมือนกับตัวยึดตำแหน่งส่วนหัว บวกด้วย:

| Placeholder | Description |
|-------------|-------------|
| `{page}` | Page number (respects format setting) / หมายเลขหน้า |

### Examples / ตัวอย่าง

```toml
[footer]
left = ""
center = "{page}"
right = ""
skip_cover = true
```

```toml
[footer]
left = "© 2024 My Company"
center = "{page}"
right = "{date}"
```

---

## Complete Example Configurations {#ch05-examples}

### Basic English Document / เอกสารภาษาอังกฤษพื้นฐาน

```toml
# md2docx.toml - Basic configuration

[document]
title = "Software User Manual"
author = "Technical Team"
date = "auto"
language = "en"

[template]
file = "templates/company.docx"

[output]
file = "output/manual.docx"

[toc]
enabled = true
depth = 3

[page_numbers]
enabled = true
format = "Page {n} of {total}"

[header]
left = "{title}"
right = "{chapter}"

[footer]
center = "{page}"
```

### Thai Document / เอกสารภาษาไทย

```toml
# md2docx.toml - Thai document configuration

[document]
title = "คู่มือการใช้งานซอฟต์แวร์"
author = "ทีมเทคนิค"
date = "auto"
language = "th"

[template]
file = "templates/thai-template.docx"

[output]
file = "output/คู่มือการใช้งาน.docx"

[toc]
enabled = true
depth = 3
title = "สารบัญ"

[fonts]
default = "TH Sarabun New"
thai = "TH Sarabun New"
normal_based_size = 28  # 14pt

[page_numbers]
enabled = true
format = "หน้า {n} จาก {total}"

[header]
left = "{title}"
right = "{chapter}"

[footer]
center = "{page}"
```

### Advanced Multi-Project Setup / การตั้งค่าหลายโครงการขั้นสูง

```toml
# md2docx.toml - Advanced configuration with all options

[document]
title = "Enterprise Documentation"
subtitle = "Complete Reference Guide"
author = "Documentation Team"
date = "{year}-{month}-{day}"
language = "en"
version = "3.2.1"

[template]
dir = "./templates/enterprise/"
validate = true

[output]
file = "dist/{version}-documentation-{date}.docx"
timestamp = false

[toc]
enabled = true
depth = 4
title = "Contents"

[fonts]
default = "Segoe UI"
thai = "TH Sarabun New"
code = "JetBrains Mono"
fallback = ["Arial Unicode MS", "Tahoma"]
h1_based_size = 36
h1_based_color = "1F4E79"

[code]
theme = "light"
show_filename = true
show_line_numbers = true
highlight_lines = true
font = "JetBrains Mono"
source_root = "../src"

[chapters]
pattern = "ch*_*.md"
sort = "numeric"
skip_cover = true

[appendices]
pattern = "ap*_*.md"
prefix = "Appendix"
sort = "alphabetic"

[page_numbers]
enabled = true
skip_cover = true
skip_chapter_first = false
format = "{n} / {total}"
start_at = 1

[header]
left = "{title} v{version}"
center = ""
right = "{chapter}"
skip_cover = true

[footer]
left = "© 2024 Enterprise Corp"
center = "{page}"
right = "Confidential"
skip_cover = true
```

### Minimal Configuration / การตั้งค่าขั้นต่ำ

```toml
# md2docx.toml - Minimal configuration
# Uses all defaults except document title

[document]
title = "My Document"
```

---

## Configuration Validation {#ch05-validation}

md2docx validates your configuration file and reports errors:

md2docx ตรวจสอบไฟล์การตั้งค่าของคุณและรายงานข้อผิดพลาด:

### Common Errors / ข้อผิดพลาดทั่วไป

| Error | Cause | Solution |
|-------|-------|----------|
| `Invalid TOML syntax` | Syntax error in config file | Check for missing quotes, commas / ตรวจสอบเครื่องหมายคำพูด ลูกน้ำ |
| `Unknown option` | Typo in option name | Check spelling / ตรวจสอบการสะกด |
| `Invalid value type` | Wrong type (e.g., string vs number) | Use correct type / ใช้ประเภทที่ถูกต้อง |
| `File not found` | Template or output path invalid | Check file paths / ตรวจสอบพาธไฟล์ |

### Validation Command / คำสั่งตรวจสอบ

```bash
# Validate config without building
md2docx build --config md2docx.toml --dry-run

# Or use validate-template for template validation
md2docx validate-template template.docx
```
