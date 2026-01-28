---
title: "บทนำ"
language: th
---

# บทที่ 1 บทนำ {#ch01}

ยินดีต้อนรับสู่คู่มือการใช้งาน md2docx บทนี้จะแนะนำให้คุณรู้จักกับ md2docx ทำความเข้าใจคุณสมบัติหลัก และเริ่มต้นใช้งานได้อย่างรวดเร็ว

## md2docx คืออะไร?

**md2docx** เป็นเครื่องมือแปลงเอกสาร Markdown เป็นไฟล์ DOCX (Microsoft Word) ที่พัฒนาด้วยภาษา Rust โดยมีจุดเด่นดังนี้:

- **รองรับภาษาไทยอย่างสมบูรณ์** - แสดงผลภาษาไทยและอังกฤษได้อย่างถูกต้อง พร้อมการจัดการฟอนต์อัตโนมัติ
- **สร้าง DOCX มาตรฐาน** - ผลลัพธ์เป็นไฟล์ DOCX ที่เปิดได้ใน Microsoft Word และโปรแกรมอื่นๆ
- **ปรับแต่ง Template ได้** - ใช้ Template ของคุณเองเพื่อกำหนดรูปแบบเอกสาร
- **รองรับฟีเจอร์ครบถ้วน** - ตาราง รูปภาพ Code blocks สารบัญ และอื่นๆ
- **ใช้งานง่าย** - คำสั่ง CLI ที่เรียบง่าย หรือใช้เป็น Library ในโปรแกรมของคุณ

### ทำไมต้องใช้ md2docx?

หากคุณเคยเขียนเอกสารด้วย Markdown แล้วต้องการแปลงเป็น DOCX คุณอาจเจอปัญหาดังนี้:

| ปัญหา | วิธีแก้แบบเดิม | วิธีแก้ด้วย md2docx |
|-------|------------------|---------------------|
| ภาษาไทยแสดงผลผิด | ต้องปรับฟอนต์ทีละตัว | จัดการฟอนต์อัตโนมัติ |
| รูปแบบไม่สวยงาม | ต้องแก้ใน Word ทีละไฟล์ | ใช้ Template กำหนดรูปแบบครั้งเดียว |
| สารบัญต้องสร้างเอง | ทำด้วยมือ ลำบาก | สร้างอัตโนมัติ |
| Code blocks ไม่สวย | ต้องจัดรูปแบบเอง | รองรับ syntax highlighting |
| หลายไฟล์ต้องรวมเอง | Copy-paste ทีละไฟล์ | รวมไฟล์อัตโนมัติ |

md2docx แก้ปัญหาเหล่านี้ทั้งหมด ให้คุณเขียนเอกสารด้วย Markdown แล้วได้ DOCX ที่สวยงามพร้อมใช้งาน

## คุณสมบัติหลัก

### รองรับภาษาไทย-อังกฤษ

md2docx ออกแบบมาเพื่อรองรับภาษาไทยอย่างเต็มรูปแบบ:

- **ตรวจจับภาษาอัตโนมัติ** - แยกข้อความภาษาไทยและอังกฤษเพื่อใช้ฟอนต์ที่เหมาะสม
- **ฟอนต์ไทยมาตรฐาน** - รองรับ TH Sarabun New, Angsana New, Tahoma และอื่นๆ
- **การตัดคำ** - รองรับการตัดคำภาษาไทย (Thai word segmentation)
- **ตัวเลขไทย** - รองรับทั้งตัวเลขอารบิก (123) และตัวเลขไทย (๑๒๓)

ตัวอย่าง:

```markdown
# การเริ่มต้น (Getting Started)

This software supports both English and ภาษาไทย in the same document.

รองรับการแสดงผลภาษาไทยแบบเต็มรูปแบบ (Full Thai rendering support)
```

### สร้างเอกสาร DOCX มาตรฐาน

md2docx สร้างไฟล์ DOCX ที่เป็นมาตรฐาน OOXML (Office Open XML):

- **เปิดได้ใน Word** - Microsoft Word 2007 ขึ้นไป
- **เปิดได้ในโปรแกรมอื่น** - LibreOffice, Google Docs, Pages
- **รักษารูปแบบ** - การจัดรูปแบบทั้งหมดถูกเก็บไว้ในไฟล์
- **ขนาดเล็ก** - ใช้การบีบอัด ZIP ที่มีประสิทธิภาพ

### ปรับแต่ง Template ได้

คุณสามารถสร้าง Template ของคุณเองเพื่อกำหนดรูปแบบเอกสาร:

- **Styles ทั้งหมด** - กำหนดฟอนต์ สี ขนาด สำหรับแต่ละส่วนของเอกสาร
- **Header/Footer** - กำหนดหัวและท้ายกระดาษ
- **หมายเลขหน้า** - กำหนดรูปแบบและตำแหน่ง
- **Auto-update** - Styles ที่สร้างจะมีคุณสมบัติ Auto-update อัตโนมัติ

> **เคล็ดลับ:** ใช้คำสั่ง `md2docx dump-template` เพื่อสร้าง Template เริ่มต้น แล้วปรับแต่งใน Microsoft Word

### รองรับ Code Blocks

md2docx รองรับการแสดงโค้ดอย่างสวยงาม:

- **Syntax highlighting** - เน้นสีตามภาษาโปรแกรม
- **ชื่อไฟล์** - แสดงชื่อไฟล์ด้านบน Code block
- **เน้นบรรทัด** - ไฮไลต์บรรทัดที่ต้องการ
- **ฟอนต์ Monospace** - ใช้ฟอนต์ Consolas หรือฟอนต์ที่คุณกำหนด

ตัวอย่าง:

```rust,filename=main.rs,hl=3
fn main() {
    let greeting = "สวัสดีชาวโลก!";
    println!("{}", greeting);  // บรรทัดนี้ถูกเน้น
}
```

### ตารางและรูปภาพ

md2docx รองรับตารางและรูปภาพอย่างครบถ้วน:

- **ตาราง** - รองรับ Markdown tables พร้อมการจัดแนวและการผสานเซลล์
- **รูปภาพ** - รองรับ PNG, JPG, GIF, SVG (แปลงเป็น PNG อัตโนมัติ)
- **Caption** - สร้างคำอธิบายรูปภาพและตารางอัตโนมัติ
- **Figure numbering** - ตัวเลขรูปภาพอัตโนมัติ (เช่น "รูปที่ 1.2")

### คุณสมบัติอื่นๆ

- **สารบัญอัตโนมัติ** - สร้าง Table of Contents จากหัวข้อทั้งหมด
- **หมายเลขหน้า** - กำหนดรูปแบบหมายเลขหน้า และข้ามหน้าปก
- **หลายไฟล์** - รวมไฟล์ Markdown หลายไฟล์เป็นเอกสารเดียว
- **Cross-references** - อ้างอิงถึงบท รูปภาพ ตาราง อัตโนมัติ
- **Footnotes** - รองรับเชิงอรรถ
- **Mermaid diagrams** - แปลง diagram เป็นรูปภาพอัตโนมัติ

## ความต้องการของระบบ

### ระบบปฏิบัติการที่รองรับ

md2docx รองรับระบบปฏิบัติการหลักทั้งหมด:

| ระบบปฏิบัติการ | เวอร์ชันขั้นต่ำ | หมายเหตุ |
|----------------|------------------|----------|
| Windows | Windows 7 ขึ้นไป | รองรับทั้ง 32-bit และ 64-bit |
| macOS | macOS 10.13 (High Sierra) ขึ้นไป | Intel และ Apple Silicon |
| Linux | ทุก distro ที่รองรับ glibc 2.17+ | Ubuntu, Debian, Fedora, etc. |

### ความต้องการฮาร์ดแวร์

- **CPU:** 1 core ขึ้นไป
- **RAM:** 512MB ขั้นต่ำ (แนะนำ 1GB สำหรับเอกสารขนาดใหญ่)
- **Disk:** 50MB สำหรับการติดตั้ง

### ซอฟต์แวร์ที่ต้องการ

- **CLI:** ไม่ต้องการซอฟต์แวร์เพิ่มเติม
- **Library:** Rust 1.75+ (หากต้องการคอมไพล์เอง)
- **Optional:**
  - Microsoft Word 2007+ (สำหรับเปิดไฟล์ DOCX)
  - Node.js (สำหรับ Mermaid rendering แบบ CLI)

## การติดตั้ง

### การติดตั้งบน Windows

#### วิธีที่ 1: ดาวน์โหลด Binary (แนะนำ)

1. ไปที่ [GitHub Releases](https://github.com/your-repo/md2docx/releases)
2. ดาวน์โหลดไฟล์ `md2docx-x86_64-pc-windows-msvc.zip`
3. แตกไฟล์ ZIP
4. ย้าย `md2docx.exe` ไปยังโฟลเดอร์ที่อยู่ใน PATH (เช่น `C:\Program Files\md2docx`)
5. เปิด Command Prompt และพิมพ์:

```bash
md2docx --version
```

#### วิธีที่ 2: ใช้ Cargo

หากคุณมี Rust ติดตั้งอยู่แล้ว:

```bash
cargo install md2docx
```

### การติดตั้งบน macOS

#### วิธีที่ 1: ใช้ Homebrew (แนะนำ)

```bash
brew tap md2docx/tap
brew install md2docx
```

#### วิธีที่ 2: ดาวน์โหลด Binary

1. ดาวน์โหลดไฟล์ `md2docx-x86_64-apple-darwin.tar.gz` (Intel) หรือ `md2docx-aarch64-apple-darwin.tar.gz` (Apple Silicon)
2. แตกไฟล์:

```bash
tar -xzf md2docx-x86_64-apple-darwin.tar.gz
```

3. ย้ายไฟล์:

```bash
sudo mv md2docx /usr/local/bin/
```

4. ตรวจสอบการติดตั้ง:

```bash
md2docx --version
```

#### วิธีที่ 3: ใช้ Cargo

```bash
cargo install md2docx
```

### การติดตั้งบน Linux

#### วิธีที่ 1: ดาวน์โหลด Binary

1. ดาวน์โหลดไฟล์ `md2docx-x86_64-unknown-linux-gnu.tar.gz`
2. แตกไฟล์:

```bash
tar -xzf md2docx-x86_64-unknown-linux-gnu.tar.gz
```

3. ย้ายไฟล์:

```bash
sudo mv md2docx /usr/local/bin/
sudo chmod +x /usr/local/bin/md2docx
```

4. ตรวจสอบการติดตั้ง:

```bash
md2docx --version
```

#### วิธีที่ 2: ใช้ Cargo

```bash
cargo install md2docx
```

#### วิธีที่ 3: ใช้ Package Manager (บาง distro)

**Ubuntu/Debian:**

```bash
# เพิ่ม repository (ถ้ามี)
sudo add-apt-repository ppa:md2docx/ppa
sudo apt update
sudo apt install md2docx
```

**Fedora:**

```bash
sudo dnf install md2docx
```

**Arch Linux:**

```bash
yay -S md2docx
```

### การติดตั้งจาก Source Code

หากคุณต้องการคอมไพล์จาก source code:

```bash
# Clone repository
git clone https://github.com/your-repo/md2docx.git
cd md2docx

# Build
cargo build --release

# ไฟล์ binary จะอยู่ที่ target/release/md2docx
sudo cp target/release/md2docx /usr/local/bin/
```

## การใช้งานเบื้องต้น

### แปลงไฟล์เดียว

วิธีง่ายที่สุดในการใช้ md2docx คือการแปลงไฟล์ Markdown เดียว:

```bash
md2docx build -i README.md -o output.docx
```

คำสั่งนี้จะ:
- อ่านไฟล์ `README.md`
- แปลงเป็น DOCX
- บันทึกเป็น `output.docx`

### แปลงโฟลเดอร์ทั้งหมด

หากคุณมีโครงสร้างโฟลเดอร์ของเอกสาร:

```
my-docs/
├── cover.md
├── ch01_introduction.md
├── ch02_installation.md
└── custom-reference.docx
```

ใช้คำสั่ง:

```bash
md2docx build -d ./my-docs/ -o manual.docx --template custom-reference.docx
```

md2docx จะ:
- ค้นหาไฟล์ `cover.md` และ `ch*_*.md` อัตโนมัติ
- เรียงลำดับไฟล์ตามหมายเลข (ch01, ch02, ...)
- ใช้ Template `custom-reference.docx`
- สร้างสารบัญอัตโนมัติ
- บันทึกเป็น `manual.docx`

### ใช้ไฟล์ Config

สร้างไฟล์ `md2docx.toml` ในโฟลเดอร์เอกสาร:

```toml
[document]
title = "คู่มือการใช้งาน"
author = "ทีมพัฒนา"
language = "th"

[template]
file = "custom-reference.docx"

[output]
file = "output/manual.docx"

[toc]
enabled = true
depth = 3

[page_numbers]
enabled = true
skip_cover = true
```

จากนั้นรัน:

```bash
md2docx build -d ./my-docs/
```

md2docx จะอ่านการตั้งค่าจาก `md2docx.toml` อัตโนมัติ

### ตัวอย่างเอกสารแรกของคุณ

สร้างไฟล์ `hello.md`:

```markdown
# สวัสดีชาวโลก

นี่คือเอกสารแรกของฉันที่สร้างด้วย md2docx

## คุณสมบัติ

- รองรับภาษาไทย
- สร้าง DOCX ได้
- ใช้งานง่าย

## ตัวอย่างโค้ด

```rust
fn main() {
    println!("สวัสดี!");
}
```

> **หมายเหตุ:** md2docx เป็นเครื่องมือที่ยอดเยี่ยม!
```

แปลงเป็น DOCX:

```bash
md2docx build -i hello.md -o hello.docx
```

เปิดไฟล์ `hello.docx` ใน Microsoft Word คุณจะเห็นเอกสารที่จัดรูปแบบสวยงาม!

## สรุปบทนำ

ในบทนี้คุณได้เรียนรู้:

- md2docx คืออะไรและทำไมต้องใช้
- คุณสมบัติหลักของ md2docx
- ความต้องการของระบบ
- วิธีการติดตั้งบน Windows, macOS, และ Linux
- การใช้งานเบื้องต้น

ในบทถัดไป เราจะเรียนรู้เกี่ยวกับ **รูปแบบการเขียน Markdown** อย่างละเอียด เพื่อให้คุณสามารถสร้างเอกสารที่สวยงามและมีประสิทธิภาพ