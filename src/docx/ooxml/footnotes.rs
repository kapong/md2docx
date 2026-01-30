//! Generate word/footnotes.xml for DOCX

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::docx::ooxml::Paragraph;
use crate::error::Result;

/// Footnotes XML generator
#[derive(Debug)]
pub struct FootnotesXml {
    footnotes: Vec<Footnote>,
    next_id: i32,
}

#[derive(Debug)]
pub struct Footnote {
    pub id: i32,
    pub content: Vec<Paragraph>,
}

impl FootnotesXml {
    pub fn new() -> Self {
        Self {
            footnotes: Vec::new(),
            next_id: 1, // IDs start at 1 (0 and -1 are reserved)
        }
    }

    /// Add a footnote and return its ID
    pub fn add_footnote(&mut self, content: Vec<Paragraph>) -> i32 {
        let id = self.next_id;
        self.footnotes.push(Footnote { id, content });
        self.next_id += 1;
        id
    }

    /// Get the number of footnotes
    pub fn len(&self) -> usize {
        self.footnotes.len()
    }

    /// Check if there are any footnotes
    pub fn is_empty(&self) -> bool {
        self.footnotes.is_empty()
    }

    /// Get a reference to the footnotes vector
    pub fn get_footnotes(&self) -> &[Footnote] {
        &self.footnotes
    }

    /// Generate XML content for word/footnotes.xml
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        let mut root = BytesStart::new("w:footnotes");
        root.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        root.push_attribute((
            "xmlns:w14",
            "http://schemas.microsoft.com/office/word/2010/wordml",
        ));
        writer.write_event(Event::Start(root))?;

        // Add separator (id -1)
        self.write_separator(&mut writer, -1, "separator")?;

        // Add continuation separator (id 0)
        self.write_separator(&mut writer, 0, "continuationSeparator")?;

        // Add user footnotes
        for footnote in &self.footnotes {
            let mut ft = BytesStart::new("w:footnote");
            ft.push_attribute(("w:id", footnote.id.to_string().as_str()));
            writer.write_event(Event::Start(ft))?;

            for p in &footnote.content {
                p.write_xml(&mut writer, None)?;
            }

            writer.write_event(Event::End(BytesEnd::new("w:footnote")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("w:footnotes")))?;
        Ok(writer.into_inner().into_inner())
    }

    fn write_separator<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        id: i32,
        type_: &str,
    ) -> Result<()> {
        let mut ft = BytesStart::new("w:footnote");
        ft.push_attribute(("w:type", type_));
        ft.push_attribute(("w:id", id.to_string().as_str()));
        writer.write_event(Event::Start(ft))?;

        writer.write_event(Event::Start(BytesStart::new("w:p")))?;
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        if id == -1 {
            writer.write_event(Event::Empty(BytesStart::new("w:separator")))?;
        } else {
            writer.write_event(Event::Empty(BytesStart::new("w:continuationSeparator")))?;
        }
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        writer.write_event(Event::End(BytesEnd::new("w:p")))?;

        writer.write_event(Event::End(BytesEnd::new("w:footnote")))?;
        Ok(())
    }
}

impl Default for FootnotesXml {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::ooxml::Run;

    #[test]
    fn test_footnotes_xml_new() {
        let footnotes = FootnotesXml::new();
        assert_eq!(footnotes.next_id, 1);
        assert!(footnotes.is_empty());
        assert_eq!(footnotes.len(), 0);
    }

    #[test]
    fn test_add_footnote() {
        let mut footnotes = FootnotesXml::new();

        let content = vec![Paragraph::new().add_text("First footnote")];
        let id1 = footnotes.add_footnote(content);

        assert_eq!(id1, 1);
        assert_eq!(footnotes.len(), 1);
        assert!(!footnotes.is_empty());
        assert_eq!(footnotes.next_id, 2);

        let content2 = vec![Paragraph::new().add_text("Second footnote")];
        let id2 = footnotes.add_footnote(content2);

        assert_eq!(id2, 2);
        assert_eq!(footnotes.len(), 2);
        assert_eq!(footnotes.next_id, 3);
    }

    #[test]
    fn test_footnotes_xml_to_xml() {
        let mut footnotes = FootnotesXml::new();

        let content = vec![
            Paragraph::new().add_text("This is a footnote with "),
            Paragraph::new().add_run(Run::new("bold text").bold()),
        ];
        footnotes.add_footnote(content);

        let xml = footnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<?xml version"));
        assert!(xml_str.contains("<w:footnotes"));
        assert!(xml_str
            .contains("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""));
        assert!(xml_str.contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""));
        assert!(xml_str.contains("<w:footnote w:type=\"separator\" w:id=\"-1\""));
        assert!(xml_str.contains("<w:footnote w:type=\"continuationSeparator\" w:id=\"0\""));
        assert!(xml_str.contains("<w:footnote w:id=\"1\""));
        assert!(xml_str.contains("This is a footnote with"));
        assert!(xml_str.contains("bold text"));
        assert!(xml_str.contains("<w:b/>"));
    }

    #[test]
    fn test_footnotes_xml_empty() {
        let footnotes = FootnotesXml::new();

        let xml = footnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Should still have separators even with no user footnotes
        assert!(xml_str.contains("<w:footnotes"));
        assert!(xml_str.contains("<w:footnote w:type=\"separator\" w:id=\"-1\""));
        assert!(xml_str.contains("<w:footnote w:type=\"continuationSeparator\" w:id=\"0\""));
    }

    #[test]
    fn test_footnotes_xml_multiple_footnotes() {
        let mut footnotes = FootnotesXml::new();

        footnotes.add_footnote(vec![Paragraph::new().add_text("First")]);
        footnotes.add_footnote(vec![Paragraph::new().add_text("Second")]);
        footnotes.add_footnote(vec![Paragraph::new().add_text("Third")]);

        let xml = footnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:footnote w:id=\"1\""));
        assert!(xml_str.contains("<w:footnote w:id=\"2\""));
        assert!(xml_str.contains("<w:footnote w:id=\"3\""));
        assert!(xml_str.contains("First"));
        assert!(xml_str.contains("Second"));
        assert!(xml_str.contains("Third"));
    }

    #[test]
    fn test_footnotes_xml_with_style() {
        let mut footnotes = FootnotesXml::new();

        let content = vec![Paragraph::with_style("FootnoteText").add_text("Styled footnote")];
        footnotes.add_footnote(content);

        let xml = footnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:pStyle w:val=\"FootnoteText\"/>"));
        assert!(xml_str.contains("Styled footnote"));
    }

    #[test]
    fn test_footnotes_xml_with_complex_formatting() {
        let mut footnotes = FootnotesXml::new();

        let content = vec![Paragraph::new()
            .add_run(Run::new("Normal "))
            .add_run(Run::new("bold").bold())
            .add_run(Run::new(" "))
            .add_run(Run::new("italic").italic())
            .add_run(Run::new(" "))
            .add_run(Run::new("red").color("FF0000"))];
        footnotes.add_footnote(content);

        let xml = footnotes.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:b/>"));
        assert!(xml_str.contains("<w:i/>"));
        assert!(xml_str.contains("<w:color w:val=\"FF0000\"/>"));
    }

    #[test]
    fn test_footnotes_xml_default() {
        let footnotes = FootnotesXml::default();

        assert_eq!(footnotes.next_id, 1);
        assert!(footnotes.is_empty());
    }
}
