//! Footer XML generator for DOCX documents
//!
//! Generates footer XML with support for:
//! - Static text
//! - Page numbers (PAGE field)
//! - Total pages (NUMPAGES field)
//! - Chapter names (STYLEREF field)
//! - Document title

use crate::docx::ooxml::header::HeaderFooterField;
use crate::error::Result;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use std::io::Cursor;

/// Footer configuration
#[derive(Debug, Clone)]
pub struct FooterConfig {
    /// Left-aligned content
    pub left: Vec<HeaderFooterField>,
    /// Center-aligned content
    pub center: Vec<HeaderFooterField>,
    /// Right-aligned content
    pub right: Vec<HeaderFooterField>,
}

impl Default for FooterConfig {
    fn default() -> Self {
        Self {
            left: vec![],
            center: vec![HeaderFooterField::PageNumber],
            right: vec![],
        }
    }
}

impl FooterConfig {
    /// Create an empty footer configuration
    pub fn empty() -> Self {
        Self {
            left: vec![],
            center: vec![],
            right: vec![],
        }
    }

    /// Check if the footer has any content
    pub fn is_empty(&self) -> bool {
        self.left.is_empty() && self.center.is_empty() && self.right.is_empty()
    }
}

/// Footer XML generator
pub struct FooterXml {
    config: FooterConfig,
    document_title: String,
}

impl FooterXml {
    /// Create a new footer XML generator
    ///
    /// # Arguments
    /// * `config` - Footer configuration with left/center/right content
    /// * `document_title` - Document title for DocumentTitle field
    pub fn new(config: FooterConfig, document_title: &str) -> Self {
        Self {
            config,
            document_title: document_title.to_string(),
        }
    }

    /// Generate footer XML bytes
    ///
    /// Returns the complete footer XML as a byte vector
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);

        // XML declaration
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with namespaces
        let mut ftr = BytesStart::new("w:ftr");
        ftr.push_attribute((
            "xmlns:w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ));
        ftr.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(ftr))?;

        // Create a paragraph with three tab stops (left, center, right)
        self.write_footer_paragraph(&mut writer)?;

        writer.write_event(Event::End(BytesEnd::new("w:ftr")))?;

        Ok(writer.into_inner().into_inner())
    }

    /// Write the footer paragraph with tab stops and content
    fn write_footer_paragraph<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:p")))?;

        // Paragraph properties with tab stops (only if there's actual content)
        let has_content = !self.config.left.is_empty()
            || !self.config.center.is_empty()
            || !self.config.right.is_empty();

        writer.write_event(Event::Start(BytesStart::new("w:pPr")))?;

        if has_content {
            // Tab stops: only add when there's content to position
            let mut tabs = Vec::new();
            if !self.config.center.is_empty() {
                tabs.push(("center", "4513"));
            }
            if !self.config.right.is_empty() {
                tabs.push(("right", "9026"));
            }
            if !tabs.is_empty() {
                writer.write_event(Event::Start(BytesStart::new("w:tabs")))?;
                for (val, pos) in &tabs {
                    let mut tab = BytesStart::new("w:tab");
                    tab.push_attribute(("w:val", *val));
                    tab.push_attribute(("w:pos", *pos));
                    writer.write_event(Event::Empty(tab))?;
                }
                writer.write_event(Event::End(BytesEnd::new("w:tabs")))?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:pPr")))?;

        // Left content
        for field in &self.config.left {
            self.write_field(writer, field)?;
        }

        // Tab to center (if center content exists)
        if !self.config.center.is_empty() {
            self.write_tab(writer)?;
            for field in &self.config.center {
                self.write_field(writer, field)?;
            }
        }

        // Tab to right (if right content exists)
        if !self.config.right.is_empty() {
            self.write_tab(writer)?;
            for field in &self.config.right {
                self.write_field(writer, field)?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("w:p")))?;
        Ok(())
    }

    /// Write a tab character
    fn write_tab<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        writer.write_event(Event::Empty(BytesStart::new("w:tab")))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;
        Ok(())
    }

    /// Write a header/footer field
    fn write_field<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        field: &HeaderFooterField,
    ) -> Result<()> {
        match field {
            HeaderFooterField::Text(text) => {
                writer.write_event(Event::Start(BytesStart::new("w:r")))?;
                let mut t = BytesStart::new("w:t");
                t.push_attribute(("xml:space", "preserve"));
                writer.write_event(Event::Start(t))?;
                writer.write_event(Event::Text(BytesText::new(text)))?;
                writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                writer.write_event(Event::End(BytesEnd::new("w:r")))?;
            }
            HeaderFooterField::DocumentTitle => {
                // Just output the title as static text
                writer.write_event(Event::Start(BytesStart::new("w:r")))?;
                let mut t = BytesStart::new("w:t");
                t.push_attribute(("xml:space", "preserve"));
                writer.write_event(Event::Start(t))?;
                writer.write_event(Event::Text(BytesText::new(&self.document_title)))?;
                writer.write_event(Event::End(BytesEnd::new("w:t")))?;
                writer.write_event(Event::End(BytesEnd::new("w:r")))?;
            }
            HeaderFooterField::PageNumber => {
                self.write_page_field(writer, "PAGE")?;
            }
            HeaderFooterField::TotalPages => {
                self.write_page_field(writer, "NUMPAGES")?;
            }
            HeaderFooterField::ChapterName => {
                self.write_styleref_field(writer)?;
            }
        }
        Ok(())
    }

    /// Write a PAGE or NUMPAGES field
    ///
    /// Word fields use the structure:
    /// - fldChar begin
    /// - instrText (field instruction)
    /// - fldChar separate
    /// - placeholder text
    /// - fldChar end
    fn write_page_field<W: std::io::Write>(
        &self,
        writer: &mut Writer<W>,
        field_type: &str,
    ) -> Result<()> {
        // Field begin
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        let mut fld_char = BytesStart::new("w:fldChar");
        fld_char.push_attribute(("w:fldCharType", "begin"));
        writer.write_event(Event::Empty(fld_char))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;

        // Field instruction
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        writer.write_event(Event::Start(BytesStart::new("w:instrText")))?;
        writer.write_event(Event::Text(BytesText::new(&format!(" {} ", field_type))))?;
        writer.write_event(Event::End(BytesEnd::new("w:instrText")))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;

        // Field separate
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        let mut fld_char = BytesStart::new("w:fldChar");
        fld_char.push_attribute(("w:fldCharType", "separate"));
        writer.write_event(Event::Empty(fld_char))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;

        // Placeholder value
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        writer.write_event(Event::Start(BytesStart::new("w:t")))?;
        writer.write_event(Event::Text(BytesText::new("1")))?;
        writer.write_event(Event::End(BytesEnd::new("w:t")))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;

        // Field end
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        let mut fld_char = BytesStart::new("w:fldChar");
        fld_char.push_attribute(("w:fldCharType", "end"));
        writer.write_event(Event::Empty(fld_char))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;

        Ok(())
    }

    /// Write STYLEREF field for chapter name (references Heading 1)
    ///
    /// The STYLEREF field automatically extracts text from the most recent
    /// paragraph with the specified style (Heading 1 for chapter titles).
    /// Uses w:fldSimple for simpler field structure.
    ///
    /// IMPORTANT: w:fldSimple is a direct child of w:p, NOT wrapped in w:r.
    fn write_styleref_field<W: std::io::Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // w:fldSimple with STYLEREF instruction - direct child of paragraph, NOT inside a run
        let mut fld_simple = BytesStart::new("w:fldSimple");
        // Use &quot; for double quotes in XML attribute
        fld_simple.push_attribute(("w:instr", "STYLEREF \"Heading 1\" \\* MERGEFORMAT"));
        writer.write_event(Event::Start(fld_simple))?;

        // Placeholder run with cached value (Word will update this)
        writer.write_event(Event::Start(BytesStart::new("w:r")))?;
        // Add w:noProof to prevent spell-checking the field result
        writer.write_event(Event::Start(BytesStart::new("w:rPr")))?;
        writer.write_event(Event::Empty(BytesStart::new("w:noProof")))?;
        writer.write_event(Event::End(BytesEnd::new("w:rPr")))?;
        writer.write_event(Event::Start(BytesStart::new("w:t")))?;
        writer.write_event(Event::Text(BytesText::new("Chapter")))?;
        writer.write_event(Event::End(BytesEnd::new("w:t")))?;
        writer.write_event(Event::End(BytesEnd::new("w:r")))?;

        writer.write_event(Event::End(BytesEnd::new("w:fldSimple")))?;

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_footer_config_default() {
        let config = FooterConfig::default();
        assert!(!config.is_empty());
        assert_eq!(config.center.len(), 1);
    }

    #[test]
    fn test_footer_config_empty() {
        let config = FooterConfig::empty();
        assert!(config.is_empty());
    }

    #[test]
    fn test_footer_xml_basic() {
        let config = FooterConfig::default();
        let footer = FooterXml::new(config, "");
        let xml = footer.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<w:ftr"));
        assert!(xml_str.contains("PAGE"));
    }

    #[test]
    fn test_footer_page_x_of_y() {
        let config = FooterConfig {
            left: vec![],
            center: vec![
                HeaderFooterField::Text("Page ".to_string()),
                HeaderFooterField::PageNumber,
                HeaderFooterField::Text(" of ".to_string()),
                HeaderFooterField::TotalPages,
            ],
            right: vec![],
        };
        let footer = FooterXml::new(config, "");
        let xml = footer.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("PAGE"));
        assert!(xml_str.contains("NUMPAGES"));
    }

    #[test]
    fn test_footer_xml_empty_config() {
        let config = FooterConfig::empty();
        let footer = FooterXml::new(config, "Test");
        let xml = footer.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // Should still generate valid XML structure but no tabs for empty config
        assert!(xml_str.contains("<w:ftr"));
        assert!(xml_str.contains("<w:p"));
        assert!(!xml_str.contains("<w:tabs"));
    }

    #[test]
    fn test_footer_xml_with_chapter_name() {
        let config = FooterConfig {
            left: vec![],
            center: vec![],
            right: vec![HeaderFooterField::ChapterName],
        };
        let footer = FooterXml::new(config, "User Manual");
        let xml = footer.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("STYLEREF"));
        assert!(xml_str.contains("Heading 1"));
    }

    #[test]
    fn test_footer_xml_all_fields() {
        let config = FooterConfig {
            left: vec![
                HeaderFooterField::DocumentTitle,
                HeaderFooterField::Text(" - ".to_string()),
            ],
            center: vec![
                HeaderFooterField::Text("Page ".to_string()),
                HeaderFooterField::PageNumber,
            ],
            right: vec![HeaderFooterField::ChapterName],
        };
        let footer = FooterXml::new(config, "My Document");
        let xml = footer.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("My Document"));
        assert!(xml_str.contains("Page "));
        assert!(xml_str.contains("PAGE"));
        assert!(xml_str.contains("STYLEREF"));
        assert!(xml_str.contains("Heading 1"));
    }
}
