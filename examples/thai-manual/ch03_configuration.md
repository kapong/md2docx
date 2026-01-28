---
title: "การตั้งค่า md2docx"
language: th
---

# บทที่ 3 การตั้งค่า md2docx {#ch03}

บทนี้จะอธิบายการตั้งค่า md2docx อย่างละเอียด ตั้งแต่ไฟล์ config พื้นฐานไปจนถึงการตั้งค่าขั้นสูง เพื่อให้คุณสามารถปรับแต่ง md2docx ให้เหมาะกับความต้องการของคุณ

## ภาพรวมการตั้งค่า

md2docx มีระบบการตั้งค่าที่ยืดหยุ่น คุณสามารถตั้งค่าได้ 3 วิธี:

1. **ไฟล์ `md2docx.toml`** - วิธีหลักที่แนะนำ
2. **Command Line Arguments** - สำหรับการทดลองหรือ override
3. **Environment Variables** - สำหรับการใช้งานอัตโนมัติ

### ลำดับความสำคัญ

การตั้งค่าจะถูกนำมาใช้ตามลำดับนี้ (ค่าหลังสุดจะ override ค่าก่อนหน้า):

1. ค่าเริ่มต้นของ md2docx
2. ไฟล์ `md2docx.toml`
3. Environment Variables
4. Command Line Arguments

## ไฟล์ md2docx.toml

### ตำแหน่งไฟล์

md2docx จะค้นหาไฟล์ `md2docx.toml` ในตำแหน่งต่อไปนี้ (ตามลำดับ):

1. ไฟล์ที่ระบุด้วย `--config` หรือ `-c`
2. โฟลเดอร์ปัจจุบัน (`./md2docx.toml`)
3. โฟลเดอร์เอกสาร (`<docs-dir>/md2docx.toml`)

### โครงสร้างไฟล์พื้นฐาน

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

## การตั้งค่าเอกสาร [document]

ส่วนนี้กำหนดข้อมูลเกี่ยวกับเอกสารโดยรวม

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `title` | string | - | ชื่อเอกสาร |
| `subtitle` | string | - | ชื่อรองเอกสาร |
| `author` | string | - | ผู้แต่ง |
| `date` | string | "auto" | วันที่ (หรือ "auto") |
| `language` | string | "th" | ภาษา ("th" หรือ "en") |

### ตัวอย่างการตั้งค่า

```toml
[document]
title = "คู่มือการใช้งานระบบ ERP"
subtitle = "เวอร์ชัน 2.0"
author = "ทีมพัฒนาซอฟต์แวร์"
date = "2025-01-26"
language = "th"
```

### คำอธิบายรายตัวเลือก

#### title

ชื่อหลักของเอกสาร จะถูกใช้ใน:
- หน้าปก
- Header ของเอกสาร
- ชื่อไฟล์ (หากไม่ระบุใน `[output]`)

```toml
[document]
title = "คู่มือการใช้งาน md2docx"
```

#### subtitle

ชื่อรองของเอกสาร แสดงใต้ชื่อหลักบนหน้าปก

```toml
[document]
title = "คู่มือการใช้งาน md2docx"
subtitle = "แปลง Markdown เป็น DOCX อย่างมืออาชีพ"
```

#### author

ชื่อผู้แต่งหรือทีมที่จัดทำเอกสาร

```toml
[document]
author = "ทีมพัฒนา md2docx"
```

#### date

วันที่ของเอกสาร สามารถใช้:
- `"auto"` - ใช้วันที่ปัจจุบันอัตโนมัติ
- `"2025-01-26"` - วันที่เฉพาะเจาะจง (รูปแบบ YYYY-MM-DD)
- `"มกราคม 2569"` - วันที่แบบภาษาไทย

```toml
[document]
date = "auto"  # หรือ "2025-01-26"
```

#### language

ภาษาของเอกสาร ส่งผลต่อ:
- ฟอนต์เริ่มต้น
- สตริงที่แปลแล้ว (เช่น "สารบัญ" vs "Table of Contents")
- การจัดรูปแบบวันที่

```toml
[document]
language = "th"  # หรือ "en"
```

## การตั้งค่า Template [template]

ส่วนนี้กำหนด Template ที่จะใช้สร้างเอกสาร

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `file` | string | - | ชื่อไฟล์ Template |
| `validate` | boolean | true | ตรวจสอบ Styles ที่จำเป็น |

### ตัวอย่างการตั้งค่า

```toml
[template]
file = "custom-reference.docx"
validate = true
```

### คำอธิบายรายตัวเลือก

#### file

ชื่อไฟล์ Template (ไฟล์ DOCX) ที่จะใช้เป็นฐาน

```toml
[template]
file = "custom-reference.docx"
```

> **เคล็ดลับ:** ใช้คำสั่ง `md2docx dump-template` เพื่อสร้าง Template เริ่มต้น

#### validate

ตรวจสอบว่า Template มี Styles ที่จำเป็นครบถ้วนหรือไม่

```toml
[template]
validate = true  # แสดงคำเตือนหากขาด Styles
```

## การตั้งค่า Output [output]

ส่วนนี้กำหนดตำแหน่งและชื่อไฟล์ผลลัพธ์

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `file` | string | "output.docx" | ชื่อไฟล์ผลลัพธ์ |
| `format` | string | "docx" | รูปแบบไฟล์ (อนาคต) |

### ตัวอย่างการตั้งค่า

```toml
[output]
file = "output/คู่มือ-md2docx.docx"
```

### คำอธิบายรายตัวเลือก

#### file

ชื่อไฟล์ผลลัพธ์ สามารถระบุ:
- เพียงชื่อไฟล์: `"output.docx"`
- พาธสัมพัทธ์: `"output/manual.docx"`
- พาธสัมบูรณ์: `"/Users/user/docs/manual.docx"`

```toml
[output]
file = "output/คู่มือ-md2docx.docx"
```

## การตั้งค่าสารบัญ [toc]

ส่วนนี้กำหนดการสร้าง Table of Contents (สารบัญ)

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `enabled` | boolean | false | เปิดใช้สารบัญ |
| `depth` | number | 3 | ระดับหัวข้อสูงสุด |
| `title` | string | - | ชื่อสารบัญ (ถ้าไม่ระบุใช้ค่าจากภาษา) |

### ตัวอย่างการตั้งค่า

```toml
[toc]
enabled = true
depth = 3
title = "สารบัญ"
```

### คำอธิบายรายตัวเลือก

#### enabled

เปิดหรือปิดการสร้างสารบัญ

```toml
[toc]
enabled = true  # สร้างสารบัญ
```

#### depth

ระดับหัวข้อสูงสุดที่จะรวมในสารบัญ

- `1` - เฉพาะหัวข้อระดับ 1 (`#`)
- `2` - หัวข้อระดับ 1-2 (`#`, `##`)
- `3` - หัวข้อระดับ 1-3 (`#`, `##`, `###`)
- `4` - หัวข้อระดับ 1-4 (`#`, `##`, `###`, `####`)

```toml
[toc]
depth = 3  # รวมหัวข้อระดับ 1-3
```

#### title

ชื่อของสารบัญ หากไม่ระบุจะใช้ค่าจากภาษา:
- ภาษาไทย: "สารบัญ"
- ภาษาอังกฤษ: "Table of Contents"

```toml
[toc]
title = "สารบัญ"  # หรือ "ดัชนี"
```

## การตั้งค่าหมายเลขหน้า [page_numbers]

ส่วนนี้กำหนดการแสดงหมายเลขหน้า

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `enabled` | boolean | false | เปิดใช้หมายเลขหน้า |
| `skip_cover` | boolean | true | ข้ามหน้าปก |
| `skip_chapter_first` | boolean | true | ข้ามหน้าแรกของแต่ละบท |
| `format` | string | "{n}" | รูปแบบหมายเลขหน้า |

### ตัวอย่างการตั้งค่า

```toml
[page_numbers]
enabled = true
skip_cover = true
skip_chapter_first = true
format = "หน้า {n}"
```

### คำอธิบายรายตัวเลือก

#### enabled

เปิดหรือปิดการแสดงหมายเลขหน้า

```toml
[page_numbers]
enabled = true  # แสดงหมายเลขหน้า
```

#### skip_cover

ข้ามการแสดงหมายเลขหน้าบนหน้าปก

```toml
[page_numbers]
skip_cover = true  # ไม่แสดงหมายเลขหน้าบนหน้าปก
```

#### skip_chapter_first

ข้ามการแสดงหมายเลขหน้าบนหน้าแรกของแต่ละบท

```toml
[page_numbers]
skip_chapter_first = true  # ไม่แสดงหมายเลขหน้าบนหน้าแรกของแต่ละบท
```

#### format

รูปแบบการแสดงหมายเลขหน้า รองรับ placeholders:
- `{n}` - หมายเลขหน้าปัจจุบัน
- `{total}` - จำนวนหน้าทั้งหมด

```toml
[page_numbers]
format = "หน้า {n}"  # แสดง "หน้า 1", "หน้า 2", ...
format = "{n} / {total}"  # แสดง "1 / 100", "2 / 100", ...
format = "{n}"  # แสดงเฉพาะหมายเลข
```

## การตั้งค่า Header [header]

ส่วนนี้กำหนดเนื้อหาในส่วนหัวของแต่ละหน้า

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `left` | string | "{title}" | ข้อความซ้าย |
| `center` | string | "" | ข้อความกลาง |
| `right` | string | "{chapter}" | ข้อความขวา |
| `skip_cover` | boolean | true | ข้ามหน้าปก |

### ตัวอย่างการตั้งค่า

```toml
[header]
left = "{title}"
center = ""
right = "{chapter}"
skip_cover = true
```

### คำอธิบายรายตัวเลือก

#### left, center, right

ข้อความที่จะแสดงในส่วนหัว รองรับ placeholders:
- `{title}` - ชื่อเอกสาร
- `{chapter}` - ชื่อบทปัจจุบัน
- `{author}` - ผู้แต่ง
- `{date}` - วันที่

```toml
[header]
left = "{title}"  # แสดงชื่อเอกสารทางซ้าย
center = ""  # ไม่แสดงอะไรตรงกลาง
right = "{chapter}"  # แสดงชื่อบททางขวา
```

#### skip_cover

ข้ามการแสดง header บนหน้าปก

```toml
[header]
skip_cover = true  # ไม่แสดง header บนหน้าปก
```

## การตั้งค่า Footer [footer]

ส่วนนี้กำหนดเนื้อหาในส่วนท้ายของแต่ละหน้า

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `left` | string | "" | ข้อความซ้าย |
| `center` | string | "{page}" | ข้อความกลาง |
| `right` | string | "" | ข้อความขวา |
| `skip_cover` | boolean | true | ข้ามหน้าปก |

### ตัวอย่างการตั้งค่า

```toml
[footer]
left = ""
center = "{page}"
right = ""
skip_cover = true
```

### คำอธิบายรายตัวเลือก

#### left, center, right

ข้อความที่จะแสดงในส่วนท้าย รองรับ placeholders:
- `{page}` - หมายเลขหน้า
- `{title}` - ชื่อเอกสาร
- `{chapter}` - ชื่อบทปัจจุบัน

```toml
[footer]
left = ""  # ไม่แสดงอะไรทางซ้าย
center = "{page}"  # แสดงหมายเลขหน้าตรงกลาง
right = ""  # ไม่แสดงอะไรทางขวา
```

## การตั้งค่าฟอนต์ [fonts]

ส่วนนี้กำหนดฟอนต์ที่ใช้ในเอกสาร

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `default` | string | "TH Sarabun New" | ฟอนต์เริ่มต้น |
| `thai` | string | "TH Sarabun New" | ฟอนต์ภาษาไทย |
| `code` | string | "Consolas" | ฟอนต์โค้ด |
| `fallback` | array | [] | ฟอนต์สำรอง |

### ตัวอย่างการตั้งค่า

```toml
[fonts]
default = "TH Sarabun New"
thai = "TH Sarabun New"
code = "Consolas"
fallback = ["Arial Unicode MS", "Noto Sans Thai"]
```

### คำอธิบายรายตัวเลือก

#### default

ฟอนต์เริ่มต้นสำหรับข้อความทั่วไป

```toml
[fonts]
default = "TH Sarabun New"  # ฟอนต์ไทยมาตรฐาน
```

ฟอนต์ไทยที่แนะนำ:
- `TH Sarabun New` - ฟอนต์มาตรฐานจาก Google
- `Angsana New` - ฟอนต์คลาสสิก
- `Tahoma` - ฟอนต์ที่มีใน Windows

#### thai

ฟอนต์สำหรับข้อความภาษาไทย

```toml
[fonts]
thai = "TH Sarabun New"
```

#### code

ฟอนต์สำหรับ Code blocks

```toml
[fonts]
code = "Consolas"  # หรือ "Courier New"
```

ฟอนต์โค้ดที่แนะนำ:
- `Consolas` - ฟอนต์มาตรฐาน Windows
- `Courier New` - ฟอนต์คลาสสิก
- `Fira Code` - ฟอนต์ที่มี ligatures

#### fallback

ฟอนต์สำรอง ใช้เมื่อฟอนต์หลักไม่รองรับอักขระบางตัว

```toml
[fonts]
fallback = ["Arial Unicode MS", "Noto Sans Thai"]
```

## การตั้งค่า Code Blocks [code]

ส่วนนี้กำหนดรูปแบบการแสดง Code blocks

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `theme` | string | "light" | ธีมสี |
| `show_filename` | boolean | true | แสดงชื่อไฟล์ |
| `show_line_numbers` | boolean | false | แสดงเลขบรรทัด |
| `font` | string | "Consolas" | ฟอนต์โค้ด |

### ตัวอย่างการตั้งค่า

```toml
[code]
theme = "light"
show_filename = true
show_line_numbers = false
font = "Consolas"
```

### คำอธิบายรายตัวเลือก

#### theme

ธีมสีสำหรับ syntax highlighting

```toml
[code]
theme = "light"  # หรือ "dark"
```

- `light` - ธีมสีอ่อน (เหมาะสำหรับพิมพ์)
- `dark` - ธีมสีเข้ม (เหมาะสำหรับจอภาพ)

#### show_filename

แสดงชื่อไฟล์ด้านบน Code block

```toml
[code]
show_filename = true  # แสดงชื่อไฟล์
```

#### show_line_numbers

แสดงเลขบรรทัดทางซ้ายของ Code block

```toml
[code]
show_line_numbers = false  # ไม่แสดงเลขบรรทัด
```

#### font

ฟอนต์สำหรับ Code blocks

```toml
[code]
font = "Consolas"
```

## การตั้งค่ารูปภาพ [images]

ส่วนนี้กำหนดรูปแบบการแสดงรูปภาพ

### ตัวเลือกทั้งหมด

| ตัวเลือก | ประเภท | ค่าเริ่มต้น | คำอธิบาย |
|---------|--------|-------------|-----------|
| `max_width` | string | "100%" | ความกว้างสูงสุด |
| `default_dpi` | number | 150 | DPI เริ่มต้น |
| `figure_prefix` | string | "รูปที่" | คำนำหน้าเลขรูป |
| `auto_caption` | boolean | true | สร้าง caption อัตโนมัติ |

### ตัวอย่างการตั้งค่า

```toml
[images]
max_width = "100%"
default_dpi = 150
figure_prefix = "รูปที่"
auto_caption = true
```

### คำอธิบายรายตัวเลือก

#### max_width

ความกว้างสูงสุดของรูปภาพ

```toml
[images]
max_width = "100%"  # หรือ "800px"
```

- `"100%"` - ใช้ความกว้างเต็มหน้ากระดาษ
- `"80%"` - ใช้ 80% ของความกว้างหน้ากระดาษ
- `"800px"` - ใช้ความกว้าง 800 pixels

#### default_dpi

DPI (Dots Per Inch) เริ่มต้นสำหรับรูปภาพ

```toml
[images]
default_dpi = 150  # หรือ 300 สำหรับคุณภาพสูง
```

#### figure_prefix

คำนำหน้าเลขรูปภาพ

```toml
[images]
figure_prefix = "รูปที่"  # หรือ "Figure"
```

ผลลัพธ์: "รูปที่ 1.1", "รูปที่ 1.2", ...

#### auto_caption

สร้าง caption จาก alt text อัตโนมัติ

```toml
[images]
auto_caption = true  # ใช้ alt text เป็น caption
```

## ตัวอย่างการตั้งค่าแบบต่างๆ

### ตัวอย่างที่ 1: สำหรับรายงาน

```toml
[document]
title = "รายงานประจำปี 2568"
author = "แผนกบัญชี"
date = "2025-01-26"
language = "th"

[template]
file = "report-template.docx"

[output]
file = "output/รายงาน-2568.docx"

[toc]
enabled = true
depth = 2
title = "สารบัญ"

[page_numbers]
enabled = true
skip_cover = true
format = "หน้า {n}"

[header]
left = "{title}"
center = ""
right = ""
skip_cover = true

[footer]
left = ""
center = "{page}"
right = ""
skip_cover = true

[fonts]
default = "TH Sarabun New"
code = "Consolas"

[code]
theme = "light"
show_filename = false
show_line_numbers = false

[images]
max_width = "80%"
figure_prefix = "รูปที่"
auto_caption = true
```

### ตัวอย่างที่ 2: สำหรับคู่มือ

```toml
[document]
title = "คู่มือการใช้งานระบบ"
subtitle = "เวอร์ชัน 2.0"
author = "ทีมพัฒนา"
date = "auto"
language = "th"

[template]
file = "manual-template.docx"

[output]
file = "output/คู่มือ.docx"

[toc]
enabled = true
depth = 3
title = "สารบัญ"

[page_numbers]
enabled = true
skip_cover = true
skip_chapter_first = true
format = "หน้า {n}"

[header]
left = "{title}"
center = ""
right = "{chapter}"
skip_cover = true

[footer]
left = ""
center = "{page}"
right = ""
skip_cover = true

[fonts]
default = "TH Sarabun New"
thai = "TH Sarabun New"
code = "Consolas"

[code]
theme = "light"
show_filename = true
show_line_numbers = false
font = "Consolas"

[images]
max_width = "100%"
default_dpi = 150
figure_prefix = "รูปที่"
auto_caption = true
```

### ตัวอย่างที่ 3: สำหรับหนังสือ

```toml
[document]
title = "การเขียนโปรแกรมด้วย Rust"
subtitle = "สำหรับผู้เริ่มต้น"
author = "สมชาย ใจดี"
date = "2025-01-26"
language = "th"

[template]
file = "book-template.docx"

[output]
file = "output/หนังสือ-Rust.docx"

[toc]
enabled = true
depth = 4
title = "สารบัญ"

[page_numbers]
enabled = true
skip_cover = true
skip_chapter_first = true
format = "{n}"

[header]
left = ""
center = "{title}"
right = ""
skip_cover = true

[footer]
left = ""
center = "{page}"
right = ""
skip_cover = true

[fonts]
default = "TH Sarabun New"
thai = "TH Sarabun New"
code = "Fira Code"

[code]
theme = "light"
show_filename = true
show_line_numbers = true
font = "Fira Code"

[images]
max_width = "100%"
default_dpi = 300
figure_prefix = "รูปที่"
auto_caption = true
```

## การใช้ Command Line Arguments

คุณสามารถ override การตั้งค่าจากไฟล์ config ด้วย command line arguments:

### ตัวอย่างการใช้งาน

```bash
# Override ชื่อเอกสาร
md2docx build -d ./docs/ --title "ชื่อใหม่"

# Override ไฟล์ output
md2docx build -d ./docs/ -o new-output.docx

# Override template
md2docx build -d ./docs/ --template new-template.docx

# เปิด/ปิดสารบัญ
md2docx build -d ./docs/ --toc
md2docx build -d ./docs/ --no-toc

# กำหนดระดับสารบัญ
md2docx build -d ./docs/ --toc-depth 2

# เปิด/ปิดหมายเลขหน้า
md2docx build -d ./docs/ --page-numbers
md2docx build -d ./docs/ --no-page-numbers
```

### ตัวเลือกทั้งหมด

| ตัวเลือก | สั้น | คำอธิบาย |
|---------|------|-----------|
| `--config` | `-c` | ระบุไฟล์ config |
| `--title` | - | ชื่อเอกสาร |
| `--author` | - | ผู้แต่ง |
| `--template` | - | ไฟล์ template |
| `--output` | `-o` | ไฟล์ output |
| `--toc` | - | เปิดสารบัญ |
| `--no-toc` | - | ปิดสารบัญ |
| `--toc-depth` | - | ระดับสารบัญ |
| `--page-numbers` | - | เปิดหมายเลขหน้า |
| `--no-page-numbers` | - | ปิดหมายเลขหน้า |

## การใช้ Environment Variables

คุณสามารถตั้งค่าผ่าน environment variables:

### ตัวอย่างการใช้งาน

```bash
# ตั้งค่าผ่าน environment variables
export MD2DOCX_TITLE="ชื่อเอกสาร"
export MD2DOCX_AUTHOR="ผู้แต่ง"
export MD2DOCX_TEMPLATE="template.docx"

# รัน md2docx
md2docx build -d ./docs/
```

### ตัวแปรทั้งหมด

| ตัวแปร | คำอธิบาย |
|---------|-----------|
| `MD2DOCX_TITLE` | ชื่อเอกสาร |
| `MD2DOCX_AUTHOR` | ผู้แต่ง |
| `MD2DOCX_TEMPLATE` | ไฟล์ template |
| `MD2DOCX_OUTPUT` | ไฟล์ output |
| `MD2DOCX_LANGUAGE` | ภาษา (th/en) |
| `MD2DOCX_TOC` | เปิด/ปิดสารบัญ (true/false) |
| `MD2DOCX_TOC_DEPTH` | ระดับสารบัญ |
| `MD2DOCX_PAGE_NUMBERS` | เปิด/ปิดหมายเลขหน้า (true/false) |

## สรุปบทนี้

ในบทนี้คุณได้เรียนรู้การตั้งค่า md2docx อย่างละเอียด:

- **โครงสร้างไฟล์ md2docx.toml** - วิธีสร้างและจัดรูปแบบ
- **การตั้งค่าเอกสาร [document]** - ชื่อ ผู้แต่ง วันที่ ภาษา
- **การตั้งค่า Template [template]** - ไฟล์ template และการตรวจสอบ
- **การตั้งค่า Output [output]** - ตำแหน่งและชื่อไฟล์
- **การตั้งค่าสารบัญ [toc]** - การสร้างและระดับ
- **การตั้งค่าหมายเลขหน้า [page_numbers]** - รูปแบบและการข้าม
- **การตั้งค่า Header/Footer** - เนื้อหาและ placeholders
- **การตั้งค่าฟอนต์ [fonts]** - ฟอนต์ต่างๆ
- **การตั้งค่า Code blocks [code]** - ธีม ชื่อไฟล์ เลขบรรทัด
- **การตั้งค่ารูปภาพ [images]** - ขนาด DPI caption
- **การใช้ Command Line Arguments** - Override การตั้งค่า
- **การใช้ Environment Variables** - ตั้งค่าผ่าน environment

ในบทถัดไป เราจะเรียนรู้เกี่ยวกับ **การใช้งานขั้นสูง** รวมถึงการสร้าง Template การใช้งาน CLI และการแก้ไขปัญหา