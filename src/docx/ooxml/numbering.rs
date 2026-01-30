//! Numbering XML generation for lists
//!
//! This module generates the `word/numbering.xml` file that defines
//! how ordered and unordered lists are formatted in Word.

use crate::docx::builder::NumberingContext;
use crate::error::Result;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

/// Generate numbering.xml content with dynamic list instances
///
/// This creates:
/// - abstractNumId 1: Ordered list (1, 2, 3...)
/// - abstractNumId 2: Unordered list (bullet •)
///
/// Each list in the document gets its own unique numId that references
/// the appropriate abstractNumId (1 for ordered, 2 for unordered).
/// This ensures each list restarts numbering independently.
pub(crate) fn generate_numbering_xml_with_context(
    numbering_ctx: &NumberingContext,
) -> Result<Vec<u8>> {
    let buffer = Cursor::new(Vec::new());
    let mut writer = Writer::new(buffer);

    // XML declaration
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Root element with namespace
    let mut root = BytesStart::new("w:numbering");
    root.push_attribute((
        "xmlns:w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ));
    root.push_attribute((
        "xmlns:r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ));
    writer.write_event(Event::Start(root))?;

    // Abstract numbering 1: Ordered list (decimal)
    write_abstract_num_ordered(&mut writer, 1)?;

    // Abstract numbering 2: Unordered list (bullet)
    write_abstract_num_bullet(&mut writer, 2)?;

    // Generate a <w:num> for each list in the document
    // Each numId references abstractNumId 1 (ordered) or 2 (unordered)
    for list_info in &numbering_ctx.lists {
        let abstract_num_id = if list_info.is_ordered { 1 } else { 2 };
        write_num(&mut writer, list_info.num_id, abstract_num_id)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:numbering")))?;

    Ok(writer.into_inner().into_inner())
}

/// Generate numbering.xml content (legacy, for backwards compatibility)
///
/// This creates the numbering definitions for:
/// - abstractNumId 1: Ordered list (1, 2, 3...)
/// - abstractNumId 2: Unordered list (bullet •)
///
/// numId 1 references abstractNumId 1 (ordered)
/// numId 2 references abstractNumId 2 (unordered)
#[allow(dead_code)]
pub fn generate_numbering_xml() -> Result<Vec<u8>> {
    let buffer = Cursor::new(Vec::new());
    let mut writer = Writer::new(buffer);

    // XML declaration
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    // Root element with namespace
    let mut root = BytesStart::new("w:numbering");
    root.push_attribute((
        "xmlns:w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ));
    root.push_attribute((
        "xmlns:r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ));
    writer.write_event(Event::Start(root))?;

    // Abstract numbering 1: Ordered list (decimal)
    write_abstract_num_ordered(&mut writer, 1)?;

    // Abstract numbering 2: Unordered list (bullet)
    write_abstract_num_bullet(&mut writer, 2)?;

    // Num 1 references abstract 1 (ordered)
    write_num(&mut writer, 1, 1)?;

    // Num 2 references abstract 2 (unordered/bullet)
    write_num(&mut writer, 2, 2)?;

    writer.write_event(Event::End(BytesEnd::new("w:numbering")))?;

    Ok(writer.into_inner().into_inner())
}

/// Write abstract numbering definition for ordered lists
fn write_abstract_num_ordered<W: std::io::Write>(writer: &mut Writer<W>, id: u32) -> Result<()> {
    let mut elem = BytesStart::new("w:abstractNum");
    elem.push_attribute(("w:abstractNumId", id.to_string().as_str()));
    writer.write_event(Event::Start(elem))?;

    // Multi-level type
    let mut mlt = BytesStart::new("w:multiLevelType");
    mlt.push_attribute(("w:val", "hybridMultilevel"));
    writer.write_event(Event::Empty(mlt))?;

    // NSID (Numbering Style Identifier) - generates a random-looking ID for uniqueness
    let mut nsid = BytesStart::new("w:nsid");
    nsid.push_attribute((
        "w:val",
        format!("{:08X}", id.wrapping_mul(0x12345678)).as_str(),
    ));
    writer.write_event(Event::Empty(nsid))?;

    // Template link (optional, but common in Word documents)
    let mut tmpl = BytesStart::new("w:tmpl");
    tmpl.push_attribute((
        "w:val",
        format!("{:08X}", id.wrapping_mul(0x87654321)).as_str(),
    ));
    writer.write_event(Event::Empty(tmpl))?;

    // Define levels 0-8 for nesting
    for ilvl in 0..9u32 {
        write_ordered_level(writer, ilvl)?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:abstractNum")))?;
    Ok(())
}

/// Write a single level for ordered list
fn write_ordered_level<W: std::io::Write>(writer: &mut Writer<W>, ilvl: u32) -> Result<()> {
    let mut lvl = BytesStart::new("w:lvl");
    lvl.push_attribute(("w:ilvl", ilvl.to_string().as_str()));
    writer.write_event(Event::Start(lvl))?;

    // Start at 1
    let mut start = BytesStart::new("w:start");
    start.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(start))?;

    // Number format: decimal
    let mut fmt = BytesStart::new("w:numFmt");
    fmt.push_attribute(("w:val", "decimal"));
    writer.write_event(Event::Empty(fmt))?;

    // Level text: "%1" for level 0, "%2" for level 1, etc. (without the dot since we add suffix)
    let lvl_text = format!("%{}", ilvl + 1);
    let mut lt = BytesStart::new("w:lvlText");
    lt.push_attribute(("w:val", lvl_text.as_str()));
    writer.write_event(Event::Empty(lt))?;

    // Left justify
    let mut jc = BytesStart::new("w:lvlJc");
    jc.push_attribute(("w:val", "left"));
    writer.write_event(Event::Empty(jc))?;

    // Suffix: tab character after the number (Word standard)
    let mut suff = BytesStart::new("w:suff");
    suff.push_attribute(("w:val", "tab"));
    writer.write_event(Event::Empty(suff))?;

    // Paragraph properties (indentation)
    writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

    let indent_left = (ilvl + 1) * 720; // 720 twips = 0.5 inch per level
    let hanging = 360; // Hanging indent for number

    let mut ind = BytesStart::new("w:ind");
    ind.push_attribute(("w:left", indent_left.to_string().as_str()));
    ind.push_attribute(("w:hanging", hanging.to_string().as_str()));
    writer.write_event(Event::Empty(ind))?;

    writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;

    writer.write_event(Event::End(BytesEnd::new("w:lvl")))?;
    Ok(())
}

/// Write abstract numbering definition for bullet lists
fn write_abstract_num_bullet<W: std::io::Write>(writer: &mut Writer<W>, id: u32) -> Result<()> {
    let mut elem = BytesStart::new("w:abstractNum");
    elem.push_attribute(("w:abstractNumId", id.to_string().as_str()));
    writer.write_event(Event::Start(elem))?;

    // Multi-level type
    let mut mlt = BytesStart::new("w:multiLevelType");
    mlt.push_attribute(("w:val", "hybridMultilevel"));
    writer.write_event(Event::Empty(mlt))?;

    // NSID (Numbering Style Identifier) - generates a random-looking ID for uniqueness
    let mut nsid = BytesStart::new("w:nsid");
    nsid.push_attribute((
        "w:val",
        format!("{:08X}", id.wrapping_mul(0xABCDEF01)).as_str(),
    ));
    writer.write_event(Event::Empty(nsid))?;

    // Template link (optional, but common in Word documents)
    let mut tmpl = BytesStart::new("w:tmpl");
    tmpl.push_attribute((
        "w:val",
        format!("{:08X}", id.wrapping_mul(0x1FEDCBA0)).as_str(),
    ));
    writer.write_event(Event::Empty(tmpl))?;

    // Define levels 0-8 for nesting
    let bullets = ["•", "○", "▪", "•", "○", "▪", "•", "○", "▪"];
    for ilvl in 0..9u32 {
        write_bullet_level(writer, ilvl, bullets[ilvl as usize])?;
    }

    writer.write_event(Event::End(BytesEnd::new("w:abstractNum")))?;
    Ok(())
}

/// Write a single level for bullet list
fn write_bullet_level<W: std::io::Write>(
    writer: &mut Writer<W>,
    ilvl: u32,
    bullet: &str,
) -> Result<()> {
    let mut lvl = BytesStart::new("w:lvl");
    lvl.push_attribute(("w:ilvl", ilvl.to_string().as_str()));
    writer.write_event(Event::Start(lvl))?;

    // Start at 1 (required but not used for bullets)
    let mut start = BytesStart::new("w:start");
    start.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(start))?;

    // Number format: bullet
    let mut fmt = BytesStart::new("w:numFmt");
    fmt.push_attribute(("w:val", "bullet"));
    writer.write_event(Event::Empty(fmt))?;

    // Level text: the bullet character
    let mut lt = BytesStart::new("w:lvlText");
    lt.push_attribute(("w:val", bullet));
    writer.write_event(Event::Empty(lt))?;

    // Left justify
    let mut jc = BytesStart::new("w:lvlJc");
    jc.push_attribute(("w:val", "left"));
    writer.write_event(Event::Empty(jc))?;

    // Suffix: tab character after the bullet (Word standard)
    let mut suff = BytesStart::new("w:suff");
    suff.push_attribute(("w:val", "tab"));
    writer.write_event(Event::Empty(suff))?;

    // Paragraph properties (indentation)
    writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

    let indent_left = (ilvl + 1) * 720; // 720 twips = 0.5 inch per level
    let hanging = 360; // Hanging indent for bullet

    let mut ind = BytesStart::new("w:ind");
    ind.push_attribute(("w:left", indent_left.to_string().as_str()));
    ind.push_attribute(("w:hanging", hanging.to_string().as_str()));
    writer.write_event(Event::Empty(ind))?;

    writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;

    // Run properties (font for bullet)
    writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;

    // Use Symbol font for standard bullets
    let mut fonts = BytesStart::new("w:rFonts");
    fonts.push_attribute(("w:ascii", "Symbol"));
    fonts.push_attribute(("w:hAnsi", "Symbol"));
    fonts.push_attribute(("w:hint", "default"));
    writer.write_event(Event::Empty(fonts))?;

    writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;

    writer.write_event(Event::End(BytesEnd::new("w:lvl")))?;
    Ok(())
}

/// Write num element that references an abstract num with level overrides
///
/// Each w:num gets a w:lvlOverride with w:startOverride to force
/// the list to restart at 1, regardless of other lists using the
/// same abstractNumId.
fn write_num<W: std::io::Write>(
    writer: &mut Writer<W>,
    num_id: u32,
    abstract_num_id: u32,
) -> Result<()> {
    let mut num = BytesStart::new("w:num");
    num.push_attribute(("w:numId", num_id.to_string().as_str()));
    writer.write_event(Event::Start(num))?;

    // Reference the abstract numbering definition
    let mut abstract_ref = BytesStart::new("w:abstractNumId");
    abstract_ref.push_attribute(("w:val", abstract_num_id.to_string().as_str()));
    writer.write_event(Event::Empty(abstract_ref))?;

    // Add level override to force restart at 1 for level 0
    // This is crucial - without it, Word continues numbering from the previous list
    // that uses the same abstractNumId
    let mut lvl_override = BytesStart::new("w:lvlOverride");
    lvl_override.push_attribute(("w:ilvl", "0"));
    writer.write_event(Event::Start(lvl_override))?;

    let mut start_override = BytesStart::new("w:startOverride");
    start_override.push_attribute(("w:val", "1"));
    writer.write_event(Event::Empty(start_override))?;

    writer.write_event(Event::End(BytesEnd::new("w:lvlOverride")))?;

    writer.write_event(Event::End(BytesEnd::new("w:num")))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_generate_numbering_xml() {
        let xml = generate_numbering_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Check structure
        assert!(xml_str.contains("<?xml version"));
        assert!(xml_str.contains("<w:numbering"));
        assert!(xml_str.contains("<w:abstractNum"));
        assert!(xml_str.contains("w:abstractNumId=\"1\""));
        assert!(xml_str.contains("w:abstractNumId=\"2\""));
        assert!(xml_str.contains("<w:num"));
        assert!(xml_str.contains("w:numId=\"1\""));
        assert!(xml_str.contains("w:numId=\"2\""));

        // Check ordered list format
        assert!(xml_str.contains("w:val=\"decimal\""));

        // Check bullet format
        assert!(xml_str.contains("w:val=\"bullet\""));
    }
}
