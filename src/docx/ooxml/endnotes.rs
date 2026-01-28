//! Generate word/endnotes.xml for DOCX
//!
//! Endnotes are similar to footnotes but appear at the end of the document.
//! Even if no user endnotes are used, Word requires the endnotes.xml file
//! with separator entries (IDs -1 and 0) when referenced in settings.xml.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;

/// Endnotes XML generator
#[derive(Debug, Default)]
pub struct EndnotesXml {
    // We don't store user endnotes yet - just generate the separators
    // Future: Add endnote storage similar to FootnotesXml
}

impl EndnotesXml {
    pub fn new() -> Self {
        Self::default()
    }

    /// Generate XML content for word/endnotes.xml
    /// This includes the required separator entries (IDs -1 and 0)
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("w:endnotes");
        root.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        writer.write_event(Event::Start(root))?;

        // Add separator (id -1) - required by Word
        self.write_separator(&mut writer, -1, "separator")?;

        // Add continuation separator (id 0) - required by Word
        self.write_separator(&mut writer, 0, "continuationSeparator")?;

        writer.write_event(Event::End(BytesEnd::new("w:endnotes")))?;
        Ok(writer.into_inner().into_inner())
    }

    fn write_separator<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        id: i32,
        type_: &str,
    ) -> Result<()> {
        let mut en = BytesStart::new("w:endnote");
        en.push_attribute(("w:type", type_));
        en.push_attribute(("w:id", id.to_string().as_str()));
        writer.write_event(Event::Start(en))?;

        writer.write_event(Event::Start(BytesStart::new("w:p")))?;
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        if id == -1 {
            writer.write_event(Event::Empty(BytesStart::new("w:separator")))?;
        } else {
            writer.write_event(Event::Empty(BytesStart::new("w:continuationSeparator")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        writer.write_event(Event::End(BytesEnd::new("w:p")))?;

        writer.write_event(Event::End(BytesEnd::new("w:endnote")))?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_endnotes_xml_new() {
        let endnotes = EndnotesXml::new();
        let xml = endnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<?xml version"));
        assert!(xml_str.contains("<w:endnotes"));
        assert!(xml_str
            .contains("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""));
    }

    #[test]
    fn test_endnotes_xml_has_separators() {
        let endnotes = EndnotesXml::new();
        let xml = endnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Must have separator (id -1) and continuation separator (id 0)
        assert!(xml_str.contains("<w:endnote w:type=\"separator\" w:id=\"-1\""));
        assert!(xml_str.contains("<w:endnote w:type=\"continuationSeparator\" w:id=\"0\""));
        assert!(xml_str.contains("<w:separator/>"));
        assert!(xml_str.contains("<w:continuationSeparator/>"));
    }

    #[test]
    fn test_endnotes_xml_structure() {
        let endnotes = EndnotesXml::new();
        let xml = endnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Each separator should have proper paragraph structure
        assert!(xml_str.contains("<w:p><w:r><w:separator/></w:r></w:p>"));
        assert!(xml_str.contains("<w:p><w:r><w:continuationSeparator/></w:r></w:p>"));
    }
}
