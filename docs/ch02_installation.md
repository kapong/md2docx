---
title: "Installation / การติดตั้ง"
language: th
---

# Chapter 2: Installation / บทที่ 2: การติดตั้ง {#ch02}

This chapter covers how to install md2docx on Windows, macOS, and Linux.

บทนี้ครอบคลุมวิธีการติดตั้ง md2docx บน Windows, macOS และ Linux

## System Requirements / ความต้องการของระบบ

### Supported Operating Systems / ระบบปฏิบัติการที่รองรับ

| OS / ระบบปฏิบัติการ | Minimum Version / เวอร์ชันขั้นต่ำ | Notes / หมายเหตุ |
|---------------------|----------------------------------|-----------------|
| Windows | Windows 7+ | 32-bit and 64-bit |
| macOS | macOS 10.13 (High Sierra)+ | Intel and Apple Silicon |
| Linux | glibc 2.17+ | Ubuntu, Debian, Fedora, etc. |

### Hardware Requirements / ความต้องการฮาร์ดแวร์

- **CPU:** 1 core minimum
- **RAM:** 512MB minimum (1GB recommended for large documents)
- **Disk:** 50MB for installation
- **CPU:** อย่างน้อย 1 core
- **RAM:** อย่างน้อย 512MB (แนะนำ 1GB สำหรับเอกสารขนาดใหญ่)
- **Disk:** 50MB สำหรับการติดตั้ง

## Installation Methods / วิธีการติดตั้ง

### Method 1: Using Cargo (Recommended) / วิธีที่ 1: ใช้ Cargo (แนะนำ)

If you have Rust installed, this is the easiest method:

หากคุณติดตั้ง Rust แล้ว นี่คือวิธีที่ง่ายที่สุด:

```bash
cargo install md2docx
```

Verify the installation:

ตรวจสอบการติดตั้ง:

```bash
md2docx --version
```

### Method 2: Build from Source / วิธีที่ 2: Build จาก Source

Clone the repository and build:

Clone repository และ build:

```bash
# Clone the repository / Clone repository
git clone https://github.com/kapong/md2docx.git
cd md2docx

# Build in release mode / Build ในโหมด release
cargo build --release

# The binary is at target/release/md2docx
# ไฟล์ binary อยู่ที่ target/release/md2docx

# Optional: Copy to your PATH / ตัวเลือก: คัดลอกไปยัง PATH
sudo cp target/release/md2docx /usr/local/bin/
```

### Method 3: Download Binary (Coming Soon) / วิธีที่ 3: ดาวน์โหลด Binary (เร็วๆ นี้)

Pre-built binaries will be available on the GitHub releases page.

Binary ที่ build ไว้แล้วจะมีให้ดาวน์โหลดบนหน้า GitHub releases

## Platform-Specific Instructions / คำแนะนำเฉพาะแพลตฟอร์ม

### Windows

1. Install Rust using rustup:
   ติดตั้ง Rust โดยใช้ rustup:

```powershell
# Download and run the installer from https://rustup.rs
# ดาวน์โหลดและรัน installer จาก https://rustup.rs
```

2. Open a new terminal and run:
   เปิด terminal ใหม่และรัน:

```powershell
cargo install md2docx
```

3. Verify installation:
   ตรวจสอบการติดตั้ง:

```powershell
md2docx --version
```

### macOS

1. Install Rust (if not already installed):
   ติดตั้ง Rust (ถ้ายังไม่ได้ติดตั้ง):

```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
```

2. Install md2docx:
   ติดตั้ง md2docx:

```bash
cargo install md2docx
```

3. Verify installation:
   ตรวจสอบการติดตั้ง:

```bash
md2docx --version
```

### Linux

1. Install Rust:
   ติดตั้ง Rust:

```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
source $HOME/.cargo/env
```

2. Install md2docx:
   ติดตั้ง md2docx:

```bash
cargo install md2docx
```

3. Verify installation:
   ตรวจสอบการติดตั้ง:

```bash
md2docx --version
```

## Optional Dependencies / ส่วนเสริมเพิ่มเติม

### LaTeX Toolchain (for Math Equation Images)

If you want to render math equations as high-quality images (the default `renderer = "image"` mode), install a LaTeX distribution:

หากต้องการแสดงผลสมการคณิตศาสตร์เป็นภาพคุณภาพสูง (โหมด `renderer = "image"` ซึ่งเป็นค่าเริ่มต้น) ให้ติดตั้ง LaTeX:

**macOS:**

```bash
brew install --cask basictex
# or full: brew install --cask mactex
```

**Linux (Ubuntu/Debian):**

```bash
sudo apt install texlive-base texlive-extra-utils
```

**Windows:**
Download and install [MiKTeX](https://miktex.org/) or [TeX Live](https://tug.org/texlive/).

Verify the toolchain is available:

ตรวจสอบว่าเครื่องมือพร้อมใช้งาน:

```bash
dvisvgm --version
tectonic --version  # or: latex --version
```

> **Note:** If the LaTeX toolchain is not installed, md2docx will automatically fall back to native Word equation rendering (OMML). You can also explicitly set `renderer = "omml"` in your `md2docx.toml` to skip the LaTeX requirement.
>
> **หมายเหตุ:** หากไม่ได้ติดตั้ง LaTeX md2docx จะใช้การแสดงผลสมการแบบ Word ดั้งเดิม (OMML) แทนโดยอัตโนมัติ คุณสามารถตั้ง `renderer = "omml"` ใน `md2docx.toml` เพื่อข้ามความต้องการ LaTeX ได้

### Embedded Tectonic (Optional Feature)

You can build md2docx with the tectonic TeX engine compiled directly into the binary, eliminating the need to install `tectonic` CLI separately. Only `dvisvgm` is still needed externally.

สามารถ build md2docx โดยรวม tectonic TeX engine เข้าไปในไบนารีโดยตรง ไม่ต้องติดตั้ง `tectonic` CLI แยกต่างหาก ต้องใช้เฉพาะ `dvisvgm` จากภายนอกเท่านั้น

```bash
cargo install md2docx --features tectonic-lib
```

> **Note:** The `tectonic-lib` feature requires system C/C++ libraries (`harfbuzz`, `icu`, `freetype`, `graphite2`) and increases binary size significantly. On first use, tectonic downloads a ~30 MB TeX package bundle from the internet.
>
> **หมายเหตุ:** ฟีเจอร์ `tectonic-lib` ต้องใช้ไลบรารี C/C++ ของระบบ (`harfbuzz`, `icu`, `freetype`, `graphite2`) และทำให้ไบนารีใหญ่ขึ้นอย่างมาก ในการใช้งานครั้งแรก tectonic จะดาวน์โหลดชุดแพ็คเกจ TeX ขนาด ~30 MB จากอินเทอร์เน็ต

---

## Verifying Installation / ตรวจสอบการติดตั้ง

After installation, verify that md2docx is working correctly:

หลังติดตั้ง ตรวจสอบว่า md2docx ทำงานถูกต้อง:

```bash
# Check version / ตรวจสอบเวอร์ชัน
md2docx --version

# View help / ดูความช่วยเหลือ
md2docx --help

# Quick test / ทดสอบเร็ว
echo "# Hello World" | md2docx build -i - -o test.docx
```

## Updating md2docx / อัปเดต md2docx

To update to the latest version:

เพื่ออัปเดตเป็นเวอร์ชันล่าสุด:

```bash
cargo install md2docx --force
```

## Uninstalling / ถอนการติดตั้ง

To uninstall md2docx:

เพื่อถอนการติดตั้ง md2docx:

```bash
cargo uninstall md2docx
```

## Next Steps / ขั้นตอนถัดไป

Now that md2docx is installed, proceed to Chapter 3 to create your first document.

ตอนนี้ md2docx ติดตั้งแล้ว ดำเนินต่อไปยังบทที่ 3 เพื่อสร้างเอกสารแรกของคุณ
