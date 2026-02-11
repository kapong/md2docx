# Appendix B: Version History {#ap02}

This appendix documents the version history and changes made to md2docx.

ภาคผนวกนี้บันทึกประวัติเวอร์ชันและการเปลี่ยนแปลงที่ทำกับ md2docx

---

## Version 0.1.3 (2026-02-11) {#ap02-v013}

**Bug Fix Release / เวอร์ชันแก้ไขข้อบกพร่อง**

### New Features / ฟีเจอร์ใหม่

- Added `[{PAGENUM}](#bookmark)` syntax to insert Word PAGEREF field that resolves to the page number of the referenced heading/anchor
- Fixed cross-references and PAGEREF not working inside table cells

### Bug Fixes / การแก้ไขข้อบกพร่อง

- Fixed multi-level (nested) list rendering — parent item text was lost or replaced by child list items
- Fixed tight list items with nested content losing parent text when nested list started

---

## Version 0.1.2 (2026-02-11) {#ap02-v012}

**Feature & Bug Fix Release / เวอร์ชันเพิ่มฟีเจอร์และแก้ไขข้อบกพร่อง**

### New Features / ฟีเจอร์ใหม่

- Added first-line indent (0.5 inch) to normal body text paragraphs
- Header/footer tab stop positions now computed from actual page dimensions instead of hardcoded A4 values
- Cover template text elements support 5x longer title/subtitle/author values
- Custom user-defined variables in `[document]` section available as `{{key}}` placeholders in cover templates and output filenames

### Bug Fixes / การแก้ไขข้อบกพร่อง

- Fixed panic when truncating filenames containing multi-byte Thai characters

---

## Version 0.1.1 (2026-01-30) {#ap02-v011}

**Bug Fix Release / เวอร์ชันแก้ไขข้อบกพร่อง**

### Bug Fixes / การแก้ไขข้อบกพร่อง

- Fixed "Duplicate filename" ZIP error when building documents with multiple images
- Removed 15% padding from Mermaid diagram SVG canvas for cleaner rendering
- Removed border and shadow effects from Mermaid diagrams

---

## Version 0.1.0 (2026-01-30) {#ap02-v010}

**Initial Release / เวอร์ชันแรก**

This is the first public release of md2docx, a Markdown to DOCX converter with Thai/English support.

นี่คือเวอร์ชันสาธารณะแรกของ md2docx โปรแกรมแปลง Markdown เป็น DOCX พร้อมรองรับภาษาไทย/อังกฤษ

### Features / ฟีเจอร์

#### Core Conversion / การแปลงหลัก

**Markdown to DOCX conversion** - Full CommonMark support with extensions for tables, footnotes, and more.

**Thai language support** - Optimized fonts (TH Sarabun New), proper line breaking for Thai text.

**Mixed Thai-English text** - Automatic font switching for complex scripts ensures proper rendering.

**การแปลง Markdown เป็น DOCX** - รองรับ CommonMark อย่างเต็มรูปแบบพร้อมส่วนขยายสำหรับตาราง เชิงอรรถ และอื่นๆ

**รองรับภาษาไทย** - ฟอนต์ที่ปรับให้เหมาะสม (TH Sarabun New) การตัดบรรทัดที่ถูกต้องสำหรับข้อความไทย

**ข้อความผสมไทย-อังกฤษ** - การสลับฟอนต์อัตโนมัติสำหรับสคริปต์ที่ซับซ้อนรับประกันการเรนเดอร์ที่ถูกต้อง

#### Document Structure / โครงสร้างเอกสาร

**Multi-file projects** - Combine multiple markdown files into one DOCX document.

**Cover page support** - Dedicated `cover.md` for title pages with custom formatting.

**Chapter ordering** - Automatic sorting by `ch01_`, `ch02_`, etc. pattern.

**Appendix support** - Pattern-based appendix inclusion using `ap*_*.md` naming.

**Table of Contents** - Automatic TOC generation with configurable depth (1-6 levels).

**Page numbering** - Automatic page numbers in headers/footers with customizable format.

**โครงการหลายไฟล์** - รวมไฟล์ markdown หลายไฟล์เป็น DOCX เดียว

**รองรับหน้าปก** - `cover.md` เฉพาะสำหรับหน้าชื่อเรื่องพร้อมการจัดรูปแบบกำหนดเอง

**การเรียงลำดับบท** - การเรียงลำดับอัตโนมัติตามรูปแบบ `ch01_`, `ch02_`, ฯลฯ

**รองรับภาคผนวก** - การรวมภาคผนวกตามรูปแบบโดยใช้การตั้งชื่อ `ap*_*.md`

**สารบัญ** - การสร้าง TOC อัตโนมัติพร้อมความลึกที่กำหนดได้ (1-6 ระดับ)

**หมายเลขหน้า** - หมายเลขหน้าอัตโนมัติในส่วนหัว/ท้ายพร้อมรูปแบบที่กำหนดได้

#### Content Elements / องค์ประกอบเนื้อหา

**Headings** - H1-H6 with automatic styling and outline levels for TOC.

**Paragraphs** - Full Unicode support with proper text flow and justification.

**Lists** - Ordered and unordered lists with proper numbering restart between separate lists.

**Code blocks** - Monospace font formatting with optional filename display.

**Inline code** - Proper formatting within paragraphs using distinct styling.

**Blockquotes** - Styled quote blocks with left border and indentation.

**Tables** - Full table support with header rows, column alignment, and borders.

**Images** - PNG, JPEG, GIF, BMP support with automatic sizing and captions.

**Links** - Hyperlinks with proper formatting and clickable URLs.

**Footnotes** - Automatic footnote numbering with proper placement at page bottom.

**หัวข้อ** - H1-H6 พร้อมการจัดสไตล์อัตโนมัติและระดับเค้าร่างสำหรับ TOC

**ย่อหน้า** - รองรับ Unicode อย่างเต็มรูปแบบพร้อมการไหลของข้อความและการจัดชิดขอบที่ถูกต้อง

**รายการ** - รายการเรียงลำดับและไม่เรียงลำดับพร้อมการเริ่มนับใหม่ที่ถูกต้องระหว่างรายการแยก

**บล็อกโค้ด** - การจัดรูปแบบฟอนต์ monospace พร้อมการแสดงชื่อไฟล์ทางเลือก

**โค้ดอินไลน์** - การจัดรูปแบบที่ถูกต้องภายในย่อหน้าโดยใช้การจัดสไตล์ที่แตกต่าง

**บล็อกอ้างอิง** - บล็อกอ้างอิงที่จัดสไตล์พร้อมขอบซ้ายและการเยื้อง

**ตาราง** - รองรับตารางอย่างเต็มรูปแบบพร้อมแถวส่วนหัว การจัดตำแหน่งคอลัมน์ และขอบ

**รูปภาพ** - รองรับ PNG, JPEG, GIF, BMP พร้อมการปรับขนาดอัตโนมัติและคำบรรยาย

**ลิงก์** - ไฮเปอร์ลิงก์พร้อมการจัดรูปแบบที่ถูกต้องและ URL ที่คลิกได้

**เชิงอรรถ** - การนับเชิงอรรถอัตโนมัติพร้อมการวางตำแหน่งที่ถูกต้องที่ด้านล่างหน้า

#### Mermaid Diagrams / แผนผัง Mermaid

**Native Rust rendering** - No browser or external tools required for diagram generation.

**Supported diagram types** - Flowcharts, sequence diagrams, class diagrams, state diagrams, and Gantt charts.

**SVG to PNG conversion** - Automatic conversion for Word compatibility.

**Automatic sizing** - Diagrams scale to fit page width while maintaining aspect ratio.

**การเรนเดอร์ Rust แบบเนทีฟ** - ไม่ต้องการเบราว์เซอร์หรือเครื่องมือภายนอกสำหรับการสร้างแผนผัง

**ประเภทแผนผังที่รองรับ** - Flowcharts, sequence diagrams, class diagrams, state diagrams และ Gantt charts

**การแปลง SVG เป็น PNG** - การแปลงอัตโนมัติสำหรับความเข้ากันได้กับ Word

**การปรับขนาดอัตโนมัติ** - แผนผังปรับขนาดให้พอดีกับความกว้างหน้าในขณะที่รักษาอัตราส่วนภาพ

#### Template System / ระบบแม่แบบ

**Template extraction** - Extract styles from existing DOCX files for reuse.

**Template validation** - Verify required styles are present before building.

**Template generation** - Create new templates with `dump-template` command.

**Style inheritance** - Apply template styles to generated content automatically.

**การแยกแม่แบบ** - แยกสไตล์จากไฟล์ DOCX ที่มีอยู่เพื่อนำกลับมาใช้

**การตรวจสอบแม่แบบ** - ตรวจสอบว่ามีสไตล์ที่จำเป็นอยู่ก่อนสร้าง

**การสร้างแม่แบบ** - สร้างแม่แบบใหม่ด้วยคำสั่ง `dump-template`

**การสืบทอดสไตล์** - ใช้สไตล์แม่แบบกับเนื้อหาที่สร้างโดยอัตโนมัติ

#### Include Directives / คำสั่ง Include

**Markdown includes** - Include external markdown files with `{!include:path/to/file.md}` syntax.

**Code includes** - Include source code files with `{!code:path/to/file.rs}` syntax.

**Line ranges** - Include specific lines with `{!code:file.rs:10-25}` syntax.

**Nested includes** - Recursive include resolution for complex document structures.

**การรวม Markdown** - รวมไฟล์ markdown ภายนอกด้วยไวยากรณ์ `{!include:path/to/file.md}`

**การรวมโค้ด** - รวมไฟล์ซอร์สโค้ดด้วยไวยากรณ์ `{!code:path/to/file.rs}`

**ช่วงบรรทัด** - รวมบรรทัดเฉพาะด้วยไวยากรณ์ `{!code:file.rs:10-25}`

**การรวมซ้อน** - การแก้ไข include แบบเรียกซ้ำสำหรับโครงสร้างเอกสารที่ซับซ้อน

#### CLI Features / ฟีเจอร์ CLI

**Build command** - Convert markdown to DOCX with extensive options.

**Validate command** - Check template validity before use.

**Dump-template command** - Generate new templates with customization.

**Verbose output** - Detailed logging with `-v` flag for debugging.

**Output filename variables** - Use `{{currenttime}}`, `{{version}}`, etc. in output filenames.

**คำสั่ง Build** - แปลง markdown เป็น DOCX พร้อมตัวเลือกมากมาย

**คำสั่ง Validate** - ตรวจสอบความถูกต้องของแม่แบบก่อนใช้

**คำสั่ง Dump-template** - สร้างแม่แบบใหม่พร้อมการปรับแต่ง

**เอาต์พุตละเอียด** - การบันทึกรายละเอียดด้วยแฟล็ก `-v` สำหรับการดีบัก

**ตัวแปรชื่อไฟล์เอาต์พุต** - ใช้ `{{currenttime}}`, `{{version}}`, ฯลฯ ในชื่อไฟล์เอาต์พุต

### Known Limitations / ข้อจำกัดที่ทราบ

Thai text in Mermaid diagrams does not render correctly due to upstream bug in mermaid-rs-renderer crate.

WASM bindings are not fully implemented and are stubs only.

Watch mode (`--watch`) flag exists but is not functional.

Syntax highlighting (colors) for code blocks is not yet implemented.

Bibliography and citations are parsed but not processed.

ข้อความภาษาไทยในแผนผัง Mermaid ไม่เรนเดอร์อย่างถูกต้องเนื่องจากข้อบกพร่อง upstream ใน mermaid-rs-renderer crate

WASM bindings ยังไม่ได้รับการพัฒนาเต็มรูปแบบและเป็น stubs เท่านั้น

โหมด Watch (แฟล็ก `--watch`) มีอยู่แต่ไม่ทำงาน

การไฮไลต์ไวยากรณ์ (สี) สำหรับบล็อกโค้ดยังไม่ได้รับการพัฒนา

บรรณานุกรมและการอ้างอิงถูกแยกวิเคราะห์แต่ไม่ได้ประมวลผล

---

## Development Timeline {#ap02-timeline}

### January 2026 / มกราคม 2026

| Date | Milestone |
|------|-----------|
| Jan 28 | Initial commit - basic markdown to DOCX conversion |
| Jan 28 | Add mermaid diagram support with native Rust renderer |
| Jan 29 | Fix image sizing, aspect ratio, and DPI handling |
| Jan 29 | Fix relationship ID collision for images |
| Jan 29 | Convert SVG text to paths for Word compatibility |
| Jan 29 | Add 15% padding to SVG canvas for arrow rendering |
| Jan 29 | Complete mermaid diagram rendering improvements |
| Jan 29 | Implement DOCX template extraction system |
| Jan 29 | Fix cover/TOC header footer suppression |
| Jan 29 | Add table column-specific styling |
| Jan 29 | Implement image caption generation |
| Jan 30 | Add dynamic output filename with variables |
| Jan 30 | Fix Chapter 1 page numbering (starts at page 1) |
| Jan 30 | Fix cover page SVG display |
| Jan 30 | Code quality improvements and release preparation |
| Jan 30 | Add complete bilingual user manual |
| Jan 30 | Cleanup examples, reorganize documentation |

---

## Roadmap / แผนงาน {#ap02-roadmap}

### Planned Features / ฟีเจอร์ที่วางแผนไว้

#### Short-term (v0.2.0) / ระยะสั้น

Syntax highlighting for code blocks with popular language support.

Watch mode for automatic rebuilding when source files change.

Improved template style application with full extraction.

Thai text support in Mermaid diagrams (pending upstream fix).

#### Medium-term (v0.3.0) / ระยะกลาง

WASM bindings for browser-based document generation.

Bibliography and citation support with BibTeX integration.

Index generation with automatic term collection.

PDF output option using built-in conversion.

#### Long-term / ระยะยาว

Real-time collaboration features for team editing.

Cloud service integration for storage and sharing.

Plugin system for custom extensions and formats.

GUI application for non-technical users.

---

## Contributors / ผู้มีส่วนร่วม {#ap02-contributors}

**P. Phienphanich** - Creator and maintainer

### How to Contribute / วิธีมีส่วนร่วม

We welcome contributions! Please see the GitHub repository for bug reports, feature requests, pull requests, and documentation improvements.

เรายินดีรับการมีส่วนร่วม! กรุณาดูที่ GitHub repository สำหรับรายงานข้อบกพร่อง ขอฟีเจอร์ pull requests และปรับปรุงเอกสาร

**Repository:** `https://github.com/pongpanich/md2docx`

---

## License / ใบอนุญาต {#ap02-license}

md2docx is released under the **MIT License**.

md2docx เผยแพร่ภายใต้ **MIT License**

```
MIT License

Copyright (c) 2026 P. Phienphanich

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
