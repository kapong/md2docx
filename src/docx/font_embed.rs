//! Font embedding for DOCX
//!
//! Embeds TrueType/OpenType fonts into DOCX files so they render correctly
//! on systems that don't have the fonts installed.
//!
//! Per ECMA-376, embedded fonts must be obfuscated by XOR-ing the first 32 bytes
//! with a GUID-derived key. The fonts are stored as `.odttf` files in `word/fonts/`.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use crate::error::{Error, Result};

/// Represents a font file to embed
#[derive(Debug, Clone)]
pub struct EmbeddedFont {
    /// Font family name (e.g., "TH Sarabun New")
    pub font_name: String,
    /// Font variant
    pub variant: FontVariant,
    /// The obfuscated font data
    pub data: Vec<u8>,
    /// The GUID used for obfuscation (without braces, uppercase)
    pub guid: String,
    /// Filename in the DOCX archive (e.g., "font1.odttf")
    pub filename: String,
    /// Relationship ID for fontTable.xml.rels
    pub rel_id: String,
    /// Font metrics extracted from the TTF/OTF file
    pub metrics: Option<FontMetrics>,
}

/// Font variant types
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontVariant {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}

impl FontVariant {
    /// OOXML element name for this variant
    pub fn xml_element(&self) -> &'static str {
        match self {
            FontVariant::Regular => "w:embedRegular",
            FontVariant::Bold => "w:embedBold",
            FontVariant::Italic => "w:embedItalic",
            FontVariant::BoldItalic => "w:embedBoldItalic",
        }
    }
}

/// Detects font variant from filename.
/// Returns `None` for non-standard weights (Light, Medium, Thin, etc.)
/// that don't map to the 4 OOXML embed slots.
fn detect_variant(filename: &str) -> Option<FontVariant> {
    let lower = filename.to_lowercase();

    // Check for non-standard weights first — these have no OOXML embed slot
    let non_standard = [
        "light", "medium", "thin", "extralight", "ultralight",
        "semibold", "demibold", "extrabold", "ultrabold", "black", "heavy",
        "extra-light", "ultra-light", "semi-bold", "demi-bold", "extra-bold",
        "ultra-bold",
    ];
    let stem = Path::new(filename)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(filename)
        .to_lowercase();
    // Check if stem ends with a non-standard weight suffix (after hyphen)
    if let Some(suffix) = stem.rsplit('-').next() {
        if non_standard.contains(&suffix.to_lowercase().as_str()) {
            return None;
        }
    }

    if lower.contains("bolditalic") || lower.contains("bold-italic") || lower.contains("bi") {
        Some(FontVariant::BoldItalic)
    } else if lower.contains("bold") {
        Some(FontVariant::Bold)
    } else if lower.contains("italic") || lower.contains("oblique") {
        Some(FontVariant::Italic)
    } else {
        Some(FontVariant::Regular)
    }
}

/// Detects font family name from filename by removing variant suffixes and extension
fn detect_font_name(filename: &str) -> String {
    let stem = Path::new(filename)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(filename);

    // Remove common variant suffixes
    let name = stem
        .replace("-BoldItalic", "")
        .replace("-Bold", "")
        .replace("-Italic", "")
        .replace("-Regular", "")
        .replace("-Light", "")
        .replace("-Medium", "")
        .replace("-ExtraLight", "")
        .replace("-bold", "")
        .replace("-italic", "")
        .replace("-regular", "")
        .replace("-light", "")
        .replace("-medium", "");

    // Replace hyphens with spaces for typical font naming
    name.replace('-', " ")
}

/// Generate a deterministic GUID from font name and variant
/// Format: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX (uppercase hex)
fn generate_guid(font_name: &str, variant: FontVariant) -> String {
    // Use a simple hash-based approach for deterministic GUIDs
    let input = format!("{}:{:?}", font_name, variant);
    let mut hash = [0u8; 16];
    let bytes = input.as_bytes();
    for (i, &b) in bytes.iter().enumerate() {
        hash[i % 16] ^= b;
        hash[i % 16] = hash[i % 16].wrapping_mul(31).wrapping_add(b);
    }
    // Set version 4 and variant bits
    hash[6] = (hash[6] & 0x0F) | 0x40; // version 4
    hash[8] = (hash[8] & 0x3F) | 0x80; // variant 10

    format!(
        "{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
        hash[0], hash[1], hash[2], hash[3],
        hash[4], hash[5],
        hash[6], hash[7],
        hash[8], hash[9],
        hash[10], hash[11], hash[12], hash[13], hash[14], hash[15],
    )
}

/// Font embedding permission flags from the OS/2 table fsType field.
///
/// Per OpenType spec, if bit 1 (0x0002) is set, fonts must NOT be embedded.
/// Bits 2 (0x0004) and 3 (0x0008) allow Preview&Print or Editable embedding.
/// fsType == 0x0000 means Installable (most permissive).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EmbedPermission {
    /// fsType == 0x0000: No restrictions on embedding
    Installable,
    /// fsType bit 2 (0x0004): Preview & Print embedding allowed
    PreviewAndPrint,
    /// fsType bit 3 (0x0008): Editable embedding allowed
    Editable,
    /// fsType bit 1 (0x0002): Embedding is NOT permitted
    Restricted,
}

impl EmbedPermission {
    pub fn is_embeddable(&self) -> bool {
        !matches!(self, EmbedPermission::Restricted)
    }
}

/// Font metrics extracted from TTF/OTF OS/2 and name tables.
/// Used to generate correct `fontTable.xml` entries so Word renders fonts
/// at the proper scale and associates them with the right character sets.
#[derive(Debug, Clone)]
pub struct FontMetrics {
    /// Font family name from the 'name' table (nameID=1)
    pub family_name: String,
    /// Panose classification (10 bytes as hex string, e.g. "020B0502040504020204")
    pub panose: String,
    /// OOXML font family: "roman", "swiss", "modern", "script", "decorative", or "auto"
    pub family: String,
    /// Windows charset: "00" (ANSI), "DE" (Thai), "02" (Symbol), etc.
    pub charset: String,
    /// Pitch: "fixed" or "variable"
    pub pitch: String,
    /// Unicode signature ranges for w:sig element
    pub usb0: u32,
    pub usb1: u32,
    pub usb2: u32,
    pub usb3: u32,
    pub csb0: u32,
    pub csb1: u32,
}

/// Locate a table in a TrueType/OpenType font and return its (offset, length).
fn find_table(data: &[u8], table_tag: &[u8; 4]) -> Option<(usize, usize)> {
    if data.len() < 12 {
        return None;
    }
    let num_tables = u16::from_be_bytes([data[4], data[5]]) as usize;
    for i in 0..num_tables {
        let dir_offset = 12 + i * 16;
        if dir_offset + 16 > data.len() {
            return None;
        }
        if &data[dir_offset..dir_offset + 4] == table_tag {
            let table_offset = u32::from_be_bytes([
                data[dir_offset + 8],
                data[dir_offset + 9],
                data[dir_offset + 10],
                data[dir_offset + 11],
            ]) as usize;
            let table_length = u32::from_be_bytes([
                data[dir_offset + 12],
                data[dir_offset + 13],
                data[dir_offset + 14],
                data[dir_offset + 15],
            ]) as usize;
            return Some((table_offset, table_length));
        }
    }
    None
}

/// Read the font family name from the 'name' table (nameID=1).
///
/// Tries Windows platform (platformID=3) first, then Mac (platformID=1).
fn read_font_name(data: &[u8]) -> Option<String> {
    let (table_offset, _) = find_table(data, b"name")?;
    if table_offset + 6 > data.len() {
        return None;
    }
    let count = u16::from_be_bytes([data[table_offset + 2], data[table_offset + 3]]) as usize;
    let string_offset =
        u16::from_be_bytes([data[table_offset + 4], data[table_offset + 5]]) as usize;
    let storage_start = table_offset + string_offset;

    // Each name record: platformID(2) + encodingID(2) + languageID(2) + nameID(2) + length(2) + offset(2) = 12 bytes
    let records_start = table_offset + 6;

    // First pass: try Windows platform (platformID=3, encodingID=1 = UTF-16BE)
    for i in 0..count {
        let rec = records_start + i * 12;
        if rec + 12 > data.len() {
            break;
        }
        let platform_id = u16::from_be_bytes([data[rec], data[rec + 1]]);
        let name_id = u16::from_be_bytes([data[rec + 6], data[rec + 7]]);
        if platform_id == 3 && name_id == 1 {
            let length = u16::from_be_bytes([data[rec + 8], data[rec + 9]]) as usize;
            let offset = u16::from_be_bytes([data[rec + 10], data[rec + 11]]) as usize;
            let start = storage_start + offset;
            if start + length <= data.len() && length >= 2 {
                // Decode UTF-16BE
                let utf16: Vec<u16> = data[start..start + length]
                    .chunks_exact(2)
                    .map(|c| u16::from_be_bytes([c[0], c[1]]))
                    .collect();
                return Some(String::from_utf16_lossy(&utf16));
            }
        }
    }

    // Second pass: try Mac platform (platformID=1, encodingID=0 = MacRoman)
    for i in 0..count {
        let rec = records_start + i * 12;
        if rec + 12 > data.len() {
            break;
        }
        let platform_id = u16::from_be_bytes([data[rec], data[rec + 1]]);
        let name_id = u16::from_be_bytes([data[rec + 6], data[rec + 7]]);
        if platform_id == 1 && name_id == 1 {
            let length = u16::from_be_bytes([data[rec + 8], data[rec + 9]]) as usize;
            let offset = u16::from_be_bytes([data[rec + 10], data[rec + 11]]) as usize;
            let start = storage_start + offset;
            if start + length <= data.len() {
                // MacRoman is mostly ASCII-compatible
                return Some(
                    data[start..start + length]
                        .iter()
                        .map(|&b| b as char)
                        .collect(),
                );
            }
        }
    }

    None
}

/// Read comprehensive font metrics from raw TTF/OTF file bytes.
///
/// Extracts Panose classification, family class, charset, pitch,
/// and Unicode/CodePage signature ranges from the OS/2 table,
/// plus the font family name from the name table.
pub fn read_font_metrics(data: &[u8]) -> Option<FontMetrics> {
    let (os2_offset, os2_length) = find_table(data, b"OS/2")?;

    // Need at least 58 bytes for ulUnicodeRange fields
    if os2_offset + 58 > data.len() {
        return None;
    }

    // OS/2 version at offset 0 (determines available fields)
    let version = u16::from_be_bytes([data[os2_offset], data[os2_offset + 1]]);

    // sFamilyClass at offset 30-31
    let family_class =
        i16::from_be_bytes([data[os2_offset + 30], data[os2_offset + 31]]);
    let class_id = (family_class >> 8) as i8;

    // Panose at offset 32-41 (10 bytes)
    let panose_bytes = &data[os2_offset + 32..os2_offset + 42];
    let panose: String = panose_bytes.iter().map(|b| format!("{:02X}", b)).collect();

    // OOXML family from sFamilyClass
    let family = match class_id {
        1..=5 | 7 => "roman",
        8 => "swiss",
        9 => "modern",       // Monospaced (but check Panose too)
        10 => "script",
        12 => "decorative",
        _ => "auto",
    }
    .to_string();

    // Pitch from Panose bProportion (byte at index 3)
    let pitch = if panose_bytes[3] == 9 { "fixed" } else { "variable" }.to_string();

    // ulUnicodeRange1-4 at offsets 42-57
    let usb0 = u32::from_be_bytes([
        data[os2_offset + 42], data[os2_offset + 43],
        data[os2_offset + 44], data[os2_offset + 45],
    ]);
    let usb1 = u32::from_be_bytes([
        data[os2_offset + 46], data[os2_offset + 47],
        data[os2_offset + 48], data[os2_offset + 49],
    ]);
    let usb2 = u32::from_be_bytes([
        data[os2_offset + 50], data[os2_offset + 51],
        data[os2_offset + 52], data[os2_offset + 53],
    ]);
    let usb3 = u32::from_be_bytes([
        data[os2_offset + 54], data[os2_offset + 55],
        data[os2_offset + 56], data[os2_offset + 57],
    ]);

    // ulCodePageRange1-2 at offsets 78-85 (requires OS/2 version >= 1, table length >= 86)
    let (csb0, csb1) = if version >= 1 && os2_length >= 86 && os2_offset + 86 <= data.len() {
        (
            u32::from_be_bytes([
                data[os2_offset + 78], data[os2_offset + 79],
                data[os2_offset + 80], data[os2_offset + 81],
            ]),
            u32::from_be_bytes([
                data[os2_offset + 82], data[os2_offset + 83],
                data[os2_offset + 84], data[os2_offset + 85],
            ]),
        )
    } else {
        (0, 0)
    };

    // Charset from code page ranges
    let charset = if csb0 & 0x0001_0000 != 0 {
        "DE" // Thai (codepage 874)
    } else if csb0 & 0x8000_0000 != 0 {
        "02" // Symbol
    } else if csb0 & 0x0000_0004 != 0 {
        "01" // DEFAULT (has Latin-2)
    } else {
        "00" // ANSI (Latin-1)
    }
    .to_string();

    // Font family name from 'name' table
    let family_name = read_font_name(data).unwrap_or_default();

    Some(FontMetrics {
        family_name,
        panose,
        family,
        charset,
        pitch,
        usb0,
        usb1,
        usb2,
        usb3,
        csb0,
        csb1,
    })
}

/// Read the OS/2 fsType embedding permission from raw font file bytes.
///
/// Parses the TrueType/OpenType table directory to locate the OS/2 table,
/// then reads the `fsType` field at offset 8 within that table.
pub fn read_fs_type(data: &[u8]) -> Option<u16> {
    let (table_offset, _) = find_table(data, b"OS/2")?;
    // fsType is at offset 8 within the OS/2 table
    if table_offset + 10 > data.len() {
        return None;
    }
    let fs_type = u16::from_be_bytes([data[table_offset + 8], data[table_offset + 9]]);
    Some(fs_type)
}

/// Classify embedding permission from the fsType value.
pub fn classify_embed_permission(fs_type: u16) -> EmbedPermission {
    if fs_type == 0x0000 {
        EmbedPermission::Installable
    } else if fs_type & 0x0002 != 0 {
        EmbedPermission::Restricted
    } else if fs_type & 0x0008 != 0 {
        EmbedPermission::Editable
    } else if fs_type & 0x0004 != 0 {
        EmbedPermission::PreviewAndPrint
    } else {
        // Unknown bits — treat as restricted to be safe
        EmbedPermission::Restricted
    }
}

/// Check if a font file allows embedding based on its OS/2 fsType field.
pub fn check_embed_permission(path: &Path) -> Result<EmbedPermission> {
    let data = std::fs::read(path)?;
    match read_fs_type(&data) {
        Some(fs_type) => Ok(classify_embed_permission(fs_type)),
        None => {
            // No OS/2 table found — assume embeddable (some older fonts lack it)
            Ok(EmbedPermission::Installable)
        }
    }
}

/// Obfuscate font data per ECMA-376 §15.2.12
///
/// The first 32 bytes of the font data are XOR'd with a key derived from the GUID.
/// The GUID (without hyphens) is parsed as hex bytes, reversed, and repeated to form 32 bytes.
fn obfuscate_font_data(data: &[u8], guid: &str) -> Vec<u8> {
    // Parse GUID hex digits (remove hyphens)
    let hex_str: String = guid.chars().filter(|c| *c != '-').collect();
    let key_bytes: Vec<u8> = (0..hex_str.len())
        .step_by(2)
        .filter_map(|i| u8::from_str_radix(&hex_str[i..i + 2], 16).ok())
        .collect();

    // Build 32-byte XOR key: the 16 GUID bytes reversed, then repeated
    let mut xor_key = [0u8; 32];
    for i in 0..16 {
        xor_key[i] = key_bytes[15 - i];
        xor_key[16 + i] = key_bytes[15 - i];
    }

    let mut result = data.to_vec();
    for i in 0..32.min(result.len()) {
        result[i] ^= xor_key[i];
    }
    result
}

/// Scan a directory for font files and group them by font family
pub fn scan_font_dir(dir: &Path) -> Result<HashMap<String, Vec<(PathBuf, FontVariant)>>> {
    if !dir.exists() || !dir.is_dir() {
        return Err(Error::Template(format!(
            "Font embed directory does not exist: {}",
            dir.display()
        )));
    }

    let mut families: HashMap<String, Vec<(PathBuf, FontVariant)>> = HashMap::new();

    for entry in std::fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if !path.is_file() {
            continue;
        }

        let ext = path
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();

        if ext != "ttf" && ext != "otf" {
            continue;
        }

        let filename = path
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_string();

        let font_name = detect_font_name(&filename);
        let variant = match detect_variant(&filename) {
            Some(v) => v,
            None => {
                eprintln!(
                    "Skipping non-standard weight font: {} (OOXML only supports Regular/Bold/Italic/BoldItalic)",
                    filename
                );
                continue;
            }
        };

        // Check font embedding permission via OS/2 fsType
        match check_embed_permission(&path) {
            Ok(permission) => {
                if !permission.is_embeddable() {
                    eprintln!(
                        "Skipping restricted font (embedding not permitted): {}",
                        filename
                    );
                    continue;
                }
            }
            Err(e) => {
                eprintln!(
                    "Warning: Could not read embedding permission for {}: {}",
                    filename, e
                );
                // Continue anyway — font may still be embeddable
            }
        }

        families
            .entry(font_name)
            .or_default()
            .push((path, variant));
    }

    Ok(families)
}

/// Prepare embedded fonts from a directory
///
/// Reads font files, obfuscates them, and returns `EmbeddedFont` entries
/// ready to be added to the DOCX archive.
pub fn prepare_embedded_fonts(
    dir: &Path,
    font_names: &[&str],
) -> Result<Vec<EmbeddedFont>> {
    let families = scan_font_dir(dir)?;
    let mut result = Vec::new();
    let mut font_counter = 1u32;

    for requested_name in font_names {
        // Find matching family (case-insensitive partial match)
        let requested_lower = requested_name.to_lowercase();
        let matching: Vec<_> = families
            .iter()
            .filter(|(name, _)| {
                let name_lower = name.to_lowercase();
                name_lower == requested_lower
                    || name_lower.replace(' ', "") == requested_lower.replace(' ', "")
                    || name_lower.contains(&requested_lower)
                    || requested_lower.contains(&name_lower)
            })
            .collect();

        if matching.is_empty() {
            eprintln!(
                "Warning: Font '{}' not found in {}",
                requested_name,
                dir.display()
            );
            continue;
        }

        for (family_name, variants) in &matching {
            for (path, variant) in *variants {
                let raw_data = std::fs::read(path)?;
                let metrics = read_font_metrics(&raw_data);

                // Use font name from the 'name' table if available, else from filename
                let real_name = metrics
                    .as_ref()
                    .map(|m| m.family_name.as_str())
                    .filter(|n| !n.is_empty())
                    .unwrap_or(requested_name);

                let guid = generate_guid(family_name, *variant);
                let obfuscated = obfuscate_font_data(&raw_data, &guid);
                let filename = format!("font{}.odttf", font_counter);
                let rel_id = format!("rIdFont{}", font_counter);

                result.push(EmbeddedFont {
                    font_name: real_name.to_string(),
                    variant: *variant,
                    data: obfuscated,
                    guid,
                    filename,
                    rel_id,
                    metrics,
                });

                font_counter += 1;
            }
        }
    }

    Ok(result)
}

/// Grouping of embedded fonts by font name for fontTable.xml generation
pub fn group_by_font_name(fonts: &[EmbeddedFont]) -> HashMap<String, Vec<&EmbeddedFont>> {
    let mut groups: HashMap<String, Vec<&EmbeddedFont>> = HashMap::new();
    for font in fonts {
        groups
            .entry(font.font_name.clone())
            .or_default()
            .push(font);
    }
    groups
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_variant() {
        assert_eq!(detect_variant("Font-Regular.ttf"), Some(FontVariant::Regular));
        assert_eq!(detect_variant("Font-Bold.ttf"), Some(FontVariant::Bold));
        assert_eq!(detect_variant("Font-Italic.otf"), Some(FontVariant::Italic));
        assert_eq!(
            detect_variant("Font-BoldItalic.ttf"),
            Some(FontVariant::BoldItalic)
        );
        assert_eq!(detect_variant("Font.ttf"), Some(FontVariant::Regular));
        assert_eq!(detect_variant("TUFont-bold.ttf"), Some(FontVariant::Bold));
        // Non-standard weights should return None
        assert_eq!(detect_variant("Font-Light.ttf"), None);
        assert_eq!(detect_variant("Font-Medium.otf"), None);
        assert_eq!(detect_variant("Font-Thin.ttf"), None);
        assert_eq!(detect_variant("Font-SemiBold.ttf"), None);
        assert_eq!(detect_variant("Font-ExtraLight.ttf"), None);
        assert_eq!(detect_variant("Font-Black.ttf"), None);
    }

    #[test]
    fn test_detect_font_name() {
        assert_eq!(detect_font_name("Pridi-Regular.ttf"), "Pridi");
        assert_eq!(detect_font_name("Pridi-Bold.ttf"), "Pridi");
        assert_eq!(detect_font_name("TUFont-bold.ttf"), "TUFont");
        assert_eq!(detect_font_name("TUFont.ttf"), "TUFont");
    }

    #[test]
    fn test_generate_guid() {
        let guid = generate_guid("Test Font", FontVariant::Regular);
        // Should be valid GUID format
        assert_eq!(guid.len(), 36);
        assert_eq!(guid.chars().filter(|c| *c == '-').count(), 4);
    }

    #[test]
    fn test_obfuscate_roundtrip() {
        let data = vec![0xAA; 64];
        let guid = "12345678-1234-4234-8234-123456789012";
        let obfuscated = obfuscate_font_data(&data, guid);
        // Obfuscating again with same key should restore original
        let restored = obfuscate_font_data(&obfuscated, guid);
        assert_eq!(data, restored);
    }

    #[test]
    fn test_classify_embed_permission() {
        assert_eq!(classify_embed_permission(0x0000), EmbedPermission::Installable);
        assert!(classify_embed_permission(0x0000).is_embeddable());

        assert_eq!(classify_embed_permission(0x0002), EmbedPermission::Restricted);
        assert!(!classify_embed_permission(0x0002).is_embeddable());

        assert_eq!(classify_embed_permission(0x0004), EmbedPermission::PreviewAndPrint);
        assert!(classify_embed_permission(0x0004).is_embeddable());

        assert_eq!(classify_embed_permission(0x0008), EmbedPermission::Editable);
        assert!(classify_embed_permission(0x0008).is_embeddable());

        // Restricted bit takes precedence
        assert_eq!(classify_embed_permission(0x000A), EmbedPermission::Restricted);
        assert!(!classify_embed_permission(0x000A).is_embeddable());
    }

    #[test]
    fn test_read_fs_type_debug_fonts() {
        let debug_dir = Path::new("debug");
        if !debug_dir.exists() {
            return; // skip if debug fonts not available
        }
        // Test a few known fonts from the debug directory
        let test_cases = vec![
            ("Pridi-Regular.ttf", 0x0000u16),     // Installable (SIL OFL)
            ("TUFont.ttf", 0x0008u16),             // Editable
            ("ChulabhornLikitText-Regular.otf", 0x0000u16), // Installable
        ];
        for (filename, expected_fs_type) in test_cases {
            let path = debug_dir.join(filename);
            if path.exists() {
                let data = std::fs::read(&path).unwrap();
                let fs_type = read_fs_type(&data);
                assert_eq!(
                    fs_type,
                    Some(expected_fs_type),
                    "fsType mismatch for {}",
                    filename
                );
            }
        }
    }

    #[test]
    fn test_read_font_metrics() {
        let font_dir = Path::new("docs/template/fonts");
        if !font_dir.exists() {
            return;
        }
        let path = font_dir.join("NotoSansThai-Regular.ttf");
        if !path.exists() {
            return;
        }
        let data = std::fs::read(&path).unwrap();
        let metrics = read_font_metrics(&data).expect("should read metrics");
        assert_eq!(metrics.family_name, "Noto Sans Thai");
        assert_eq!(metrics.panose, "020B0502040504020204");
        assert_eq!(metrics.charset, "DE"); // Thai
        assert_eq!(metrics.pitch, "variable");
        assert!(metrics.csb0 & 0x0001_0000 != 0, "should have Thai codepage bit set");
    }

    #[test]
    fn test_read_font_name() {
        let font_dir = Path::new("docs/template/fonts");
        if !font_dir.exists() {
            return;
        }
        let path = font_dir.join("Srisakdi-Regular.ttf");
        if !path.exists() {
            return;
        }
        let data = std::fs::read(&path).unwrap();
        let name = read_font_name(&data).expect("should read name");
        assert_eq!(name, "Srisakdi");
    }
}
