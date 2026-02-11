# CLI Reference {#ch07}

This chapter provides a complete reference for the md2docx command-line interface. All commands, options, and examples are documented here.

บทนี้ให้ข้อมูลอ้างอิงที่สมบูรณ์สำหรับอินเตอร์เฟซบรรทัดคำสั่งของ md2docx บันทึกคำสั่ง ตัวเลือก และตัวอย่างทั้งหมดไว้ที่นี่

---

## Overview {#ch07-overview}

### English

md2docx provides a command-line interface for converting Markdown to DOCX. The CLI supports:

- Single file conversion
- Directory-based project builds
- Template management
- Configuration file processing

### ภาษาไทย

md2docx ให้อินเตอร์เฟซบรรทัดคำสั่งสำหรับการแปลง Markdown เป็น DOCX CLI รองรับ:

- การแปลงไฟล์เดี่ยว
- การสร้างโครงการแบบไดเรกทอรี
- การจัดการแม่แบบ
- การประมวลผลไฟล์การตั้งค่า

### Command Structure / โครงสร้างคำสั่ง

```bash
md2docx [GLOBAL_OPTIONS] <COMMAND> [COMMAND_OPTIONS]
```

### Available Commands / คำสั่งที่มี

| Command | Description |
|---------|-------------|
| `build` | Convert markdown to DOCX / แปลง markdown เป็น DOCX |
| `help` | Show help information / แสดงข้อมูลความช่วยเหลือ |

---

## Global Options {#ch07-global-options}

### English

Global options apply to all commands and must be specified before the command name.

### ภาษาไทย

ตัวเลือกทั่วไปนำไปใช้กับทุกคำสั่งและต้องระบุก่อนชื่อคำสั่ง

### Options / ตัวเลือก

| Option | Short | Description |
|--------|-------|-------------|
| `--help` | `-h` | Show help message and exit / แสดงข้อความช่วยเหลือ |
| `--version` | `-V` | Show version information / แสดงข้อมูลเวอร์ชัน |
| `--verbose` | `-v` | Enable verbose output / เปิดใช้งานเอาต์พุตแบบละเอียด |
| `--quiet` | `-q` | Suppress non-error output / ระงับเอาต์พุตที่ไม่ใช่ข้อผิดพลาด |

### Examples / ตัวอย่าง

```bash
# Show version
md2docx --version

# Show help for specific command
md2docx build --help

# Verbose output
md2docx -v build -i input.md -o output.docx
```

---

## build Command {#ch07-build}

### English

The `build` command converts Markdown files to DOCX format. It supports both single file and directory-based builds.

### ภาษาไทย

คำสั่ง `build` แปลงไฟล์ Markdown เป็นรูปแบบ DOCX รองรับทั้งการสร้างไฟล์เดี่ยวและแบบไดเรกทอรี

### Usage / การใช้งาน

```bash
md2docx build [OPTIONS]
```

### Input Options / ตัวเลือกอินพุต

| Option | Short | Type | Description |
|--------|-------|------|-------------|
| `--input` | `-i` | string | Single input markdown file / ไฟล์ markdown อินพุตเดี่ยว |
| `--directory` | `-d` | string | Project directory to build / ไดเรกทอรีโครงการที่จะสร้าง |
| `--config` | `-c` | string | Configuration file path / พาธไฟล์การตั้งค่า |

**Note**: Use either `-i` or `-d`, not both.

**หมายเหตุ**: ใช้ `-i` หรือ `-d` อย่างใดอย่างหนึ่ง ไม่ใช่ทั้งคู่

### Output Options / ตัวเลือกเอาต์พุต

| Option | Short | Type | Default | Description |
|--------|-------|------|---------|-------------|
| `--output` | `-o` | string | `"output.docx"` | Output filename / ชื่อไฟล์เอาต์พุต |

### Template Options / ตัวเลือกแม่แบบ

| Option | Short | Type | Description |
|--------|-------|------|-------------|
| `--template` | `-t` | string | Template file path / พาธไฟล์แม่แบบ |

### TOC Options / ตัวเลือกสารบัญ

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `--toc` | boolean | `true` | Include table of contents / รวมสารบัญ |
| `--no-toc` | boolean | | Disable table of contents / ปิดใช้งานสารบัญ |
| `--toc-depth` | integer | `3` | TOC heading depth (1-6) / ความลึกของหัวข้อสารบัญ |
| `--toc-title` | string | `"Table of Contents"` | TOC title / ชื่อสารบัญ |

### Page Number Options / ตัวเลือกหมายเลขหน้า

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `--page-numbers` | boolean | `true` | Enable page numbers / เปิดใช้งานหมายเลขหน้า |
| `--no-page-numbers` | boolean | | Disable page numbers / ปิดใช้งานหมายเลขหน้า |

### Metadata Options / ตัวเลือกข้อมูลเมตา

| Option | Type | Description |
|--------|------|-------------|
| `--title` | string | Document title / ชื่อเอกสาร |
| `--author` | string | Document author / ผู้เขียนเอกสาร |
| `--date` | string | Document date / วันที่เอกสาร |
| `--language` | string | Document language (`en` or `th`) / ภาษาเอกสาร |
| `--version` | string | Document version / เวอร์ชันเอกสาร |

### Other Options / ตัวเลือกอื่นๆ

| Option | Type | Description |
|--------|------|-------------|
| `--draft` | boolean | Draft mode (skip images/TOC) / โหมดร่าง |
| `--watch` | boolean | Watch for changes and rebuild / เฝ้าดูการเปลี่ยนแปลง |

### build Examples {#ch07-build-examples}

#### Single File Conversion / การแปลงไฟล์เดี่ยว

```bash
# Basic conversion
md2docx build -i README.md -o output.docx

# With template
md2docx build -i README.md -o output.docx --template custom.docx

# With metadata
md2docx build -i README.md -o output.docx \
  --title "My Document" \
  --author "John Doe" \
  --language en
```

#### Directory Build / การสร้างไดเรกทอรี

```bash
# Build project directory
md2docx build -d ./my-docs/ -o manual.docx

# With template
md2docx build -d ./my-docs/ -o manual.docx --template company.docx

# With custom TOC
md2docx build -d ./my-docs/ -o manual.docx \
  --toc-depth 2 \
  --toc-title "Contents"
```

#### Advanced Examples / ตัวอย่างขั้นสูง

```bash
# Complete build with all options
md2docx build -d ./docs/ -o output.docx \
  --template templates/company.docx \
  --config md2docx.toml \
  --title "Software Manual" \
  --author "Tech Team" \
  --language en \
  --toc-depth 3 \
  --page-numbers

# Thai document build
md2docx build -d ./docs/ -o คู่มือ.docx \
  --template templates/thai.docx \
  --title "คู่มือการใช้งาน" \
  --author "ทีมพัฒนา" \
  --language th \
  --toc-title "สารบัญ"

# Draft mode for quick preview
md2docx build -d ./docs/ -o draft.docx --draft

# Watch mode for development
md2docx build -d ./docs/ -o output.docx --watch
```

---

## Environment Variables {#ch07-environment-variables}

### English

md2docx recognizes several environment variables for configuration. These are overridden by command-line options.

### ภาษาไทย

md2docx รู้จักตัวแปรสภาพแวดล้อมหลายตัวสำหรับการตั้งค่า สิ่งเหล่านี้จะถูกแทนที่ด้วยตัวเลือกบนบรรทัดคำสั่ง

### Variable Reference / อ้างอิงตัวแปร

| Variable | Description | Example |
|----------|-------------|---------|
| `MD2DOCX_TEMPLATE` | Default template path / พาธแม่แบบเริ่มต้น | `/path/to/template.docx` |
| `MD2DOCX_CONFIG` | Default config file / ไฟล์การตั้งค่าเริ่มต้น | `/path/to/config.toml` |
| `MD2DOCX_OUTPUT_DIR` | Default output directory / ไดเรกทอรีเอาต์พุตเริ่มต้น | `./build/` |
| `MD2DOCX_LANGUAGE` | Default language / ภาษาเริ่มต้น | `en` or `th` |
| `MD2DOCX_AUTHOR` | Default author / ผู้เขียนเริ่มต้น | `Your Name` |

### Usage Examples / ตัวอย่างการใช้งาน

```bash
# Set environment variables
export MD2DOCX_TEMPLATE="/home/user/templates/company.docx"
export MD2DOCX_LANGUAGE="th"
export MD2DOCX_AUTHOR="ทีมพัฒนา"

# Build using environment defaults
md2docx build -d ./docs/ -o output.docx

# Override with command line
md2docx build -d ./docs/ -o output.docx --language en
```

### Shell Configuration / การตั้งค่า Shell

**Bash (~/.bashrc):**

```bash
export MD2DOCX_TEMPLATE="$HOME/templates/company.docx"
export MD2DOCX_LANGUAGE="en"
```

**Zsh (~/.zshrc):**

```zsh
export MD2DOCX_TEMPLATE="$HOME/templates/company.docx"
export MD2DOCX_LANGUAGE="en"
```

**Fish (~/.config/fish/config.fish):**

```fish
set -x MD2DOCX_TEMPLATE "$HOME/templates/company.docx"
set -x MD2DOCX_LANGUAGE "en"
```

---

## Exit Codes {#ch07-exit-codes}

### English

md2docx returns specific exit codes to indicate success or failure type. Use these in scripts for error handling.

### ภาษาไทย

md2docx ส่งคืนรหัสออกเฉพาะเพื่อบ่งบอกความสำเร็จหรือประเภทความล้มเหลว ใช้สิ่งเหล่านี้ในสคริปต์สำหรับการจัดการข้อผิดพลาด

### Code Reference / อ้างอิงรหัส

| Code | Name | Description |
|------|------|-------------|
| `0` | `SUCCESS` | Operation completed successfully / การดำเนินการเสร็จสมบูรณ์ |
| `1` | `GENERAL_ERROR` | General error occurred / เกิดข้อผิดพลาดทั่วไป |
| `2` | `INVALID_INPUT` | Invalid input or arguments / อินพุตหรืออาร์กิวเมนต์ไม่ถูกต้อง |
| `3` | `FILE_NOT_FOUND` | Input file not found / ไม่พบไฟล์อินพุต |
| `4` | `PARSE_ERROR` | Markdown parsing failed / การแยกวิเคราะห์ Markdown ล้มเหลว |
| `5` | `TEMPLATE_ERROR` | Template error / ข้อผิดพลาดแม่แบบ |
| `6` | `IO_ERROR` | I/O error (permissions, disk full) / ข้อผิดพลาด I/O |
| `7` | `CONFIG_ERROR` | Configuration file error / ข้อผิดพลาดไฟล์การตั้งค่า |

### Using Exit Codes in Scripts / การใช้รหัสออกในสคริปต์

**Bash:**

```bash
md2docx build -d ./docs/ -o output.docx
case $? in
    0)
        echo "Build successful"
        ;;
    3)
        echo "Error: Input directory not found"
        exit 1
        ;;
    5)
        echo "Error: Template issue - run validate-template"
        exit 1
        ;;
    *)
        echo "Error: Build failed with code $?"
        exit 1
        ;;
esac
```

**Make:**

```makefile
build:
	@md2docx build -d ./docs/ -o output.docx || \
		(echo "Build failed"; exit 1)

ci-build:
	@md2docx build -d ./docs/ -o output.docx
	@test $$? -eq 0 || (echo "CI build failed"; exit 1)
```

---

## Common Workflows {#ch07-workflows}

### English

This section provides practical examples of common md2docx workflows.

### ภาษาไทย

ส่วนนี้ให้ตัวอย่างที่ใช้ได้จริงของเวิร์กโฟลว์ md2docx ทั่วไป

### Workflow 1: Quick Single Document / เอกสารเดี่ยวด่วน

```bash
# Convert a README to DOCX
md2docx build -i README.md -o README.docx

# Or with a title
md2docx build -i README.md -o README.docx --title "Project Documentation"
```

### Workflow 2: Project Documentation / เอกสารโครงการ

```bash
# 1. Create project structure
mkdir -p my-project/docs
mkdir -p my-project/templates

# 2. Generate template
md2docx dump-template -o my-project/templates/company.docx

# 3. Customize template in Word (manual step)
# open my-project/templates/company.docx

# 4. Create config file
cat > my-project/md2docx.toml << 'EOF'
[document]
title = "My Project Manual"
author = "Development Team"

[template]
file = "templates/company.docx"

[output]
file = "build/manual.docx"
EOF

# 5. Build documentation
cd my-project
md2docx build -d ./docs/ -o manual.docx
```

### Workflow 3: CI/CD Integration / การผสานรวม CI/CD

```bash
#!/bin/bash
# build-docs.sh - CI script

set -e

echo "Validating template..."
md2docx validate-template templates/company.docx

echo "Building documentation..."
md2docx build \
    -d ./docs/ \
    -o ./dist/manual.docx \
    --template templates/company.docx \
    --config md2docx.toml \
    --toc-depth 3

echo "Build complete: ./dist/manual.docx"
```

### Workflow 4: Multi-Language Documentation / เอกสารหลายภาษา

```bash
# Build English version
md2docx build \
    -d ./docs/en/ \
    -o ./output/manual-en.docx \
    --template templates/english.docx \
    --language en \
    --toc-title "Table of Contents"

# Build Thai version
md2docx build \
    -d ./docs/th/ \
    -o ./output/manual-th.docx \
    --template templates/thai.docx \
    --language th \
    --toc-title "สารบัญ"
```

### Workflow 5: Template Development / การพัฒนาแม่แบบ

```bash
# 1. Generate base template
md2docx dump-template -o templates/dev.docx

# 2. Edit in Word and save

# 3. Validate
md2docx validate-template templates/dev.docx

# 4. Test with sample content
echo "# Test\n\nContent" > /tmp/test.md
md2docx build -i /tmp/test.md -o /tmp/test.docx --template templates/dev.docx

# 5. Open and verify
open /tmp/test.docx
```

### Workflow 6: Watch Mode Development / การพัฒนาโหมด Watch

```bash
# Terminal 1: Watch and rebuild
md2docx build -d ./docs/ -o output.docx --watch

# Terminal 2: Edit files
vim docs/ch01_introduction.md

# Changes automatically trigger rebuild
```

---

## Quick Reference Card {#ch07-quick-reference}

### Command Summary / สรุปคำสั่ง

```bash
# Build commands
md2docx build -i <file.md> -o <output.docx>          # Single file
md2docx build -d <dir/> -o <output.docx>              # Directory
md2docx build -d <dir/> -t <template.docx>            # With template

# Template commands
md2docx dump-template -o <template.docx>              # Generate template
md2docx dump-template -o <template.docx> --lang th    # Thai template
md2docx validate-template <template.docx>             # Check template

# Help
md2docx --help                                        # General help
md2docx build --help                                  # Command help
md2docx --version                                     # Show version
```

### Common Options / ตัวเลือกทั่วไป

| Task | Command |
|------|---------|
| Set title | `--title "My Title"` |
| Set author | `--author "John Doe"` |
| Set language | `--language th` |
| Set TOC depth | `--toc-depth 2` |
| Disable TOC | `--no-toc` |
| Disable page numbers | `--no-page-numbers` |
| Use config | `--config file.toml` |
| Verbose output | `-v` or `--verbose` |

### Troubleshooting Commands / คำสั่งแก้ไขปัญหา

```bash
# Check version
md2docx --version

# Validate setup
md2docx validate-template my-template.docx

# Verbose build for debugging
md2docx -v build -d ./docs/ -o output.docx

# Test with minimal options
md2docx build -i test.md -o test.docx --no-toc --no-page-numbers
```
