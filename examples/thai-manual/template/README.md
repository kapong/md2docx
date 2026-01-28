# โครงสร้าง Template Directory

โฟลเดอร์นี้เก็บไฟล์ template สำหรับสร้างเอกสาร DOCX

## ไฟล์ใน Template

### 1. cover.docx
- หน้าปกเอกสาร
- ออกแบบใน Microsoft Word
- ใช้ placeholders:
  - `{{title}}` - ชื่อเอกสาร (จาก md2docx.toml)
  - `{{subtitle}}` - ชื่อรอง (จาก md2docx.toml)
  - `{{author}}` - ผู้จัดทำ (จาก md2docx.toml)
  - `{{date}}` - วันที่ (จาก md2docx.toml หรือ auto)
  - `{{version}}` - เวอร์ชัน (จาก md2docx.toml)
  - `{{inside}}` - เนื้อหาจาก cover.md (หลัง frontmatter)
- สามารถใส่ shapes, images, colors ได้ตามต้องการ

### 2. table.docx
- ตัวอย่างตารางสำหรับกำหนดสไตล์
- แถวที่ 1: สไตล์ header (พื้นหลังสี, ตัวหนา)
- แถวที่ 2: สไตล์แถวคี่ (odd row)
- แถวที่ 3: สไตล์แถวคู่ (even row)
- แถวที่ 4+: สไตล์คอลัมน์แรก (first column)
- Caption ด้านบน: {{table_caption_prefix}} {{table_number}}: {{table_caption}}

### 3. image.docx
- ตัวอย่างรูปภาพพร้อม caption
- Caption ด้านล่าง: {{image_caption_prefix}} {{image_number}}: {{image_caption}}

### 4. header-footer.docx
- หัวกระดาษ (header): ซ้าย {{title}}, กลางว่าง, ขวา {{chapter}}
- ท้ายกระดาษ (footer): ซ้ายว่าง, กลาง {{page}}, ขวาว่าง

## วิธีสร้าง Template

1. สร้างไฟล์ DOCX ใน Microsoft Word
2. ออกแบบตามต้องการ
3. ใช้ placeholders ในรูปแบบ {{variable}}
4. บันทึกไฟล์ในโฟลเดอร์นี้
5. ระบบจะอ่านและใช้สไตล์อัตโนมัติ
