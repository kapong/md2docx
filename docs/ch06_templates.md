# Template Customization Guide {#ch06}

This chapter explains how to create and customize DOCX templates for md2docx. Templates control the visual appearance of your documents including fonts, colors, spacing, and layout.

บทนี้อธิบายวิธีการสร้างและปรับแต่งแม่แบบ DOCX สำหรับ md2docx แม่แบบควบคุมลักษณะที่มองเห็นของเอกสารของคุณรวมถึงฟอนต์ สี ระยะห่าง และเค้าโครง

---

## What Are Templates? {#ch06-what-are-templates}

### English

A template in md2docx is a DOCX file containing predefined styles. Unlike regular documents, templates don't contain content—they contain style definitions that md2docx applies to your generated documents.

Key benefits of using templates:

- **Consistent branding** - Apply company colors, fonts, and logos
- **Professional appearance** - Pre-configured spacing and typography
- **Time savings** - No need to format documents after generation
- **Easy updates** - Change the template, regenerate all documents

### ภาษาไทย

แม่แบบใน md2docx คือไฟล์ DOCX ที่มีสไตล์ที่กำหนดไว้ล่วงหน้า ไม่เหมือนกับเอกสารทั่วไป แม่แบบไม่มีเนื้อหา—แต่มีคำนิยามสไตล์ที่ md2docx นำไปใช้กับเอกสารที่สร้างของคุณ

ประโยชน์หลักของการใช้แม่แบบ:

- **แบรนด์ที่สม่ำเสมอ** - ใช้สี ฟอนต์ และโลโก้ของบริษัท
- **ลักษณะที่เป็นมืออาชีพ** - ระยะห่างและการจัดพิมพ์ที่กำหนดค่าล่วงหน้า
- **ประหยัดเวลา** - ไม่ต้องจัดรูปแบบเอกสารหลังจากสร้าง
- **อัปเดตง่าย** - เปลี่ยนแม่แบบ สร้างเอกสารทั้งหมดใหม่

---

## Using dump-template Command {#ch06-dump-template}

### English

The easiest way to create a template is using the `dump-template` command. This generates a starter template with all required styles.

### ภาษาไทย

วิธีที่ง่ายที่สุดในการสร้างแม่แบบคือใช้คำสั่ง `dump-template` คำสั่งนี้สร้างแม่แบบเริ่มต้นที่มีสไตล์ที่จำเป็นทั้งหมด

### Basic Usage / การใช้งานพื้นฐาน

```bash
# Generate default template
md2docx dump-template -o custom-reference.docx

# Generate Thai-optimized template
md2docx dump-template -o thai-template.docx --lang th

# Generate English-optimized template
md2docx dump-template -o english-template.docx --lang en

# Generate minimal template (fewer styles)
md2docx dump-template -o minimal.docx --minimal
```

### Command Options / ตัวเลือกคำสั่ง

| Option | Short | Description |
|--------|-------|-------------|
| `--output` | `-o` | Output filename (required) |
| `--lang` | | Language preset (`en` or `th`) |
| `--minimal` | | Generate minimal template |

### Language Presets / ค่าที่ตั้งล่วงหน้าตามภาษา

**English (`--lang en`):**
- Default font: Calibri
- Heading font: Calibri Light
- Code font: Consolas
- Base size: 11pt
- Heading color: Blue (#2E74B5)

**Thai (`--lang th`):**
- Default font: TH Sarabun New
- Heading font: TH Sarabun New
- Code font: Consolas
- Base size: 14pt (better for Thai readability)
- Heading color: Blue (#2E74B5)

---

## Required Styles {#ch06-required-styles}

### English

md2docx requires certain styles to be present in your template. The `dump-template` command creates all of these automatically.

### ภาษาไทย

md2docx ต้องการให้มีสไตล์บางอย่างในแม่แบบของคุณ คำสั่ง `dump-template` สร้างสไตล์เหล่านี้ทั้งหมดโดยอัตโนมัติ

### Style Reference Table / ตารางอ้างอิงสไตล์

| Style ID | Type | Purpose | Required |
|----------|------|---------|----------|
| `Title` | Paragraph | Document title on cover | Yes |
| `Subtitle` | Paragraph | Subtitle on cover | No |
| `Heading1` | Paragraph | Chapter titles (`#`) | Yes |
| `Heading2` | Paragraph | Section titles (`##`) | Yes |
| `Heading3` | Paragraph | Subsection titles (`###`) | Yes |
| `Heading4` | Paragraph | Sub-subsection (`####`) | No |
| `Normal` | Paragraph | Body text | Yes |
| `Code` | Paragraph | Code blocks | Yes |
| `CodeChar` | Character | Inline code | Yes |
| `Quote` | Paragraph | Blockquotes | No |
| `Caption` | Paragraph | Figure/table captions | No |
| `TOC1` | Paragraph | TOC level 1 | Yes |
| `TOC2` | Paragraph | TOC level 2 | Yes |
| `TOC3` | Paragraph | TOC level 3 | Yes |
| `FootnoteText` | Paragraph | Footnote content | No |
| `Hyperlink` | Character | Links | No |
| `ListParagraph` | Paragraph | List items | Yes |
| `CodeFilename` | Paragraph | Code block filename header | No |

### Style Inheritance / การสืบทอดสไตล์

Styles inherit from parent styles:

สไตล์สืบทอดจากสไตล์หลัก:

```
Normal
├── Heading1 (based on Normal)
│   ├── Heading2
│   │   ├── Heading3
│   │   │   └── Heading4
├── Quote
├── ListParagraph
├── Caption
├── FootnoteText
└── Code
```

**Key point**: Modifying `Normal` affects all child styles.

**ประเด็นสำคัญ**: การปรับเปลี่ยน `Normal` จะส่งผลต่อสไตล์ลูกทั้งหมด

---

## Template Directory Structure {#ch06-directory-structure}

### English

For advanced customization, you can organize templates in a directory structure. This allows separate styling for different elements.

### ภาษาไทย

สำหรับการปรับแต่งขั้นสูง คุณสามารถจัดระเบียบแม่แบบในโครงสร้างไดเรกทอรี ซึ่งช่วยให้สามารถจัดรูปแบบแยกตามองค์ประกอบต่างๆ ได้

### Directory Layout / โครงสร้างไดเรกทอรี

```
templates/
└── my-company-template/
    ├── styles.docx         # Core styles (required) / สไตล์หลัก (จำเป็น)
    ├── cover.docx          # Cover page template / แม่แบบหน้าปก
    ├── header-footer.docx  # Header/footer template / แม่แบบส่วนหัว/ท้าย
    ├── image.docx          # Image styling / การจัดรูปแบบรูปภาพ
    └── table.docx          # Table styling / การจัดรูปแบบตาราง
```

### File Purposes / วัตถุประสงค์ของไฟล์

#### styles.docx (Required)

Contains all paragraph and character styles. This is the minimum required file.

มีสไตล์ย่อหน้าและอักขระทั้งหมด นี่เป็นไฟล์ที่ต้องการขั้นต่ำ

#### cover.docx (Optional)

Defines the cover page layout including:
- Background colors/images
- Title positioning
- Logo placement
- Date/author formatting

กำหนดเค้าโครงหน้าปกรวมถึง:
- สีพื้นหลัง/รูปภาพ
- ตำแหน่งชื่อเรื่อง
- ตำแหน่งโลโก้
- รูปแบบวันที่/ผู้เขียน

#### header-footer.docx (Optional)

Defines header and footer content:
- Page numbering style
- Running headers (chapter names)
- Company branding in headers

กำหนดเนื้อหาส่วนหัวและส่วนท้าย:
- รูปแบบหมายเลขหน้า
- ส่วนหัววิ่ง (ชื่อบท)
- แบรนด์บริษัทในส่วนหัว

#### image.docx (Optional)

Defines image styling:
- Default image size
- Border styles
- Caption formatting

กำหนดการจัดรูปแบบรูปภาพ:
- ขนาดรูปภาพเริ่มต้น
- รูปแบบเส้นขอบ
- การจัดรูปแบบคำบรรยาย

#### table.docx (Optional)

Defines table styling:
- Header row formatting
- Border styles
- Cell padding

กำหนดการจัดรูปแบบตาราง:
- การจัดรูปแบบแถวหัวตาราง
- รูปแบบเส้นขอบ
- ระยะห่างในเซลล์

### Configuring Template Directory / การตั้งค่าไดเรกทอรีแม่แบบ

```toml
[template]
dir = "./templates/my-company-template/"
```

---

## Customizing Styles in Word {#ch06-customizing-word}

### English

After generating a template with `dump-template`, customize it in Microsoft Word:

### ภาษาไทย

หลังจากสร้างแม่แบบด้วย `dump-template` ปรับแต่งใน Microsoft Word:

### Step-by-Step Guide / คู่มือทีละขั้นตอน

#### 1. Open the Template / เปิดแม่แบบ

```bash
# Generate template first
md2docx dump-template -o my-template.docx

# Then open in Word
open my-template.docx  # macOS
# or
start my-template.docx  # Windows
```

#### 2. Access the Styles Pane / เข้าถึงบานหน้าต่าง Styles

**Word for Windows:**
- Home tab → Styles group → click small arrow (bottom-right)
- Or press `Alt+Ctrl+Shift+S`

**Word for Mac:**
- Home tab → Styles pane → click expand button
- Or press `Command+Option+Shift+S`

**Word สำหรับ Windows:**
- แท็บ Home → กลุ่ม Styles → คลิกลูกศรเล็ก (มุมขวาล่าง)
- หรือกด `Alt+Ctrl+Shift+S`

**Word สำหรับ Mac:**
- แท็บ Home → บานหน้าต่าง Styles → คลิกปุ่มขยาย
- หรือกด `Command+Option+Shift+S`

#### 3. Modify a Style / ปรับเปลี่ยนสไตล์

1. Find the style in the Styles pane (e.g., "Heading 1")
   - ค้นหาสไตล์ในบานหน้าต่าง Styles (เช่น "Heading 1")
2. Right-click → "Modify..."
   - คลิกขวา → "Modify..."
3. Change properties:
   - เปลี่ยนคุณสมบัติ:
   - **Font**: Family, size, color, bold/italic
     - **Font**: ตระกูล ขนาด สี ตัวหนา/ตัวเอียง
   - **Paragraph**: Alignment, spacing, indentation
     - **Paragraph**: การจัดแนว ระยะห่าง การเยื้อง
   - **Border**: Lines, shading
     - **Border**: เส้น การไล่ระดับสี
4. Check "Automatically update" (recommended)
   - เลือก "Automatically update" (แนะนำ)
5. Click OK
   - คลิก OK

#### 4. Common Customizations / การปรับแต่งทั่วไป

**Change Heading Colors / เปลี่ยนสีหัวข้อ:**

```
Modify "Heading 1" → Font color → Custom → #1F4E79 (dark blue)
```

**Add Company Logo to Header / เพิ่มโลโก้บริษัทในส่วนหัว:**

```
Insert → Header → Edit Header
Insert → Pictures → Select logo
Position and resize
```

**Set Default Thai Font / ตั้งค่าฟอนต์ไทยเริ่มต้น:**

```
Modify "Normal" → Font → TH Sarabun New → Size 14
```

**Adjust Code Block Spacing / ปรับระยะห่างบล็อกโค้ด:**

```
Modify "Code" → Paragraph → Before: 6pt, After: 6pt
```

#### 5. Save the Template / บันทึกแม่แบบ

```
File → Save (or Ctrl+S / Cmd+S)
```

**Important**: Keep the `.docx` extension. md2docx reads DOCX files, not `.dotx` templates.

**สำคัญ**: เก็บนามสกุล `.docx` md2docx อ่านไฟล์ DOCX ไม่ใช่แม่แบบ `.dotx`

---

## validate-template Command {#ch06-validate-template}

### English

Before using a template, validate it to ensure all required styles are present.

### ภาษาไทย

ก่อนใช้แม่แบบ ให้ตรวจสอบเพื่อให้แน่ใจว่ามีสไตล์ที่จำเป็นทั้งหมด

### Usage / การใช้งาน

```bash
# Validate a template file
md2docx validate-template my-template.docx

# Validate with verbose output
md2docx validate-template my-template.docx --verbose
```

### Output Examples / ตัวอย่างเอาต์พุต

**Valid Template / แม่แบบที่ถูกต้อง:**

```
✓ Template validation passed

Required styles: 8/8 present
Recommended styles: 12/12 present

Template is ready to use with md2docx.
```

**Missing Styles / ขาดสไตล์:**

```
✗ Template validation failed

Missing required styles:
  - Code
  - CodeChar

Missing recommended styles:
  - Caption
  - FootnoteText

Run 'md2docx dump-template' to generate a complete template.
```

### Validation Levels / ระดับการตรวจสอบ

| Level | Styles | Impact if Missing |
|-------|--------|-------------------|
| Required | Title, Heading1-3, Normal, Code, CodeChar, TOC1-3, ListParagraph | Document may not render correctly |
| Recommended | Subtitle, Heading4, Quote, Caption, FootnoteText, Hyperlink, CodeFilename | Reduced functionality |
| Optional | Custom styles | No impact |

---

## Best Practices {#ch06-best-practices}

### English

Follow these guidelines for effective template creation.

### ภาษาไทย

ปฏิบัติตามแนวทางเหล่านี้สำหรับการสร้างแม่แบบที่มีประสิทธิภาพ

### 1. Start with dump-template / เริ่มด้วย dump-template

Always generate a base template rather than creating from scratch:

สร้างแม่แบบฐานแทนการสร้างจากศูนย์เสมอ:

```bash
# Good / ดี
md2docx dump-template -o my-template.docx
# Then customize in Word / จากนั้นปรับแต่งใน Word

# Avoid / หลีกเลี่ง
# Creating empty DOCX and adding styles manually
# สร้าง DOCX ว่างและเพิ่มสไตล์ด้วยตนเอง
```

### 2. Enable Auto-Update / เปิดใช้งาน Auto-Update

When modifying styles, enable "Automatically update":

เมื่อปรับเปลี่ยนสไตล์ ให้เปิดใช้งาน "Automatically update":

- Ensures style changes apply to all uses
- Maintains consistency throughout document
- ช่วยให้การเปลี่ยนแปลงสไตล์นำไปใช้กับทุกการใช้งาน
- รักษาความสม่ำเสมอตลอดเอกสาร

### 3. Test with Real Content / ทดสอบด้วยเนื้อหาจริง

Before finalizing a template, test with actual markdown:

ก่อนสรุปแม่แบบ ให้ทดสอบด้วย markdown จริง:

```bash
# Create test markdown
echo "# Test Heading\n\nSome **bold** text." > test.md

# Build with template
md2docx build -i test.md -o test.docx --template my-template.docx

# Open and verify
open test.docx
```

### 4. Version Your Templates / กำหนดเวอร์ชันแม่แบบของคุณ

Use semantic versioning for templates:

ใช้การกำหนดเวอร์ชันแบบ semantic สำหรับแม่แบบ:

```
templates/
├── company-template-v1.0.docx
├── company-template-v1.1.docx
└── company-template-v2.0.docx
```

### 5. Document Customizations / จดบันทึกการปรับแต่ง

Keep notes on what you changed:

เก็บบันทึกสิ่งที่คุณเปลี่ยน:

```markdown
# Template Changelog

## v1.1
- Changed Heading1 color to #1F4E79 (company blue)
- Increased Normal font size to 12pt for accessibility
- Added company logo to header

## v1.0
- Initial template from dump-template
```

### 6. Thai Font Considerations / ข้อควรพิจารณาฟอนต์ไทย

When creating Thai documents:

เมื่อสร้างเอกสารภาษาไทย:

- Use TH Sarabun New (standard government font)
  - ใช้ TH Sarabun New (ฟอนต์มาตรฐานราชการ)
- Set base size to 14pt or larger (Thai needs more space)
  - ตั้งค่าขนาดฐานเป็น 14pt หรือใหญ่กว่า (ไทยต้องการพื้นที่มากกว่า)
- Ensure complex script fonts are specified
  - ตรวจสอบให้แน่ใจว่าระบุฟอนต์สคริปต์ที่ซับซ้อน
- Test with mixed Thai-English content
  - ทดสอบด้วยเนื้อหาภาษาไทย-อังกฤษผสมกัน

### 7. Template Organization / การจัดระเบียบแม่แบบ

Organize templates by purpose:

จัดระเบียบแม่แบบตามวัตถุประสงค์:

```
templates/
├── internal/
│   ├── memo-template.docx
│   └── report-template.docx
├── external/
│   ├── proposal-template.docx
│   └── whitepaper-template.docx
├── localized/
│   ├── thai-document.docx
│   └── japanese-document.docx
└── archive/
    └── old-versions/
```

### 8. Validate Before Committing / ตรวจสอบก่อนยืนยัน

Always validate templates in CI/CD:

ตรวจสอบแม่แบบใน CI/CD เสมอ:

```yaml
# .github/workflows/docs.yml
- name: Validate Template
  run: md2docx validate-template templates/company.docx

- name: Build Documentation
  run: md2docx build -d ./docs/ --template templates/company.docx
```

---

## Quick Reference / อ้างอิงด่วน

### Creating a Template from Scratch / สร้างแม่แบบจากศูนย์

```bash
# 1. Generate base template
md2docx dump-template -o my-template.docx --lang en

# 2. Open in Word and customize
open my-template.docx

# 3. Validate
md2docx validate-template my-template.docx

# 4. Use in build
md2docx build -d ./docs/ --template my-template.docx
```

### Common Style Modifications / การปรับเปลี่ยนสไตล์ทั่วไป

| Task | How To |
|------|--------|
| Change font | Modify "Normal" style → Font |
| Change heading color | Modify "Heading1" → Font color |
| Add logo | Insert → Header → Pictures |
| Change page size | Layout → Size |
| Adjust margins | Layout → Margins |
| Set line spacing | Modify style → Paragraph → Line spacing |

### Troubleshooting / การแก้ไขปัญหา

| Problem | Solution |
|---------|----------|
| Styles not applied | Validate template; check style IDs match |
| Thai text wrong font | Ensure "cs" (complex script) font is set |
| Code blocks look wrong | Check "Code" style exists |
| TOC not styled | Verify TOC1, TOC2, TOC3 styles present |
