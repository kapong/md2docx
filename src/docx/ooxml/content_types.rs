//! Generate [Content_Types].xml for DOCX

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;

/// Content types for DOCX parts
pub(crate) struct ContentTypes {
    /// Additional content types (for images, etc.)
    extensions: Vec<(String, String)>, // (extension, content_type)
    overrides: Vec<(String, String)>, // (part_name, content_type)
}

impl ContentTypes {
    pub fn new() -> Self {
        Self {
            extensions: vec![
                ("rels".to_string(), "application/vnd.openxmlformats-package.relationships+xml".to_string()),
                ("xml".to_string(), "application/xml".to_string()),
            ],
            overrides: vec![
                ("/word/document.xml".to_string(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".to_string()),
                ("/word/styles.xml".to_string(), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml".to_string()),
                ("/word/settings.xml".to_string(), "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml".to_string()),
                ("/word/fontTable.xml".to_string(), "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml".to_string()),
                // Document properties (required for Word to not show compatibility mode)
                ("/docProps/core.xml".to_string(), "application/vnd.openxmlformats-package.core-properties+xml".to_string()),
                ("/docProps/app.xml".to_string(), "application/vnd.openxmlformats-officedocument.extended-properties+xml".to_string()),
                // Web settings and theme (required for Word compatibility)
                ("/word/webSettings.xml".to_string(), "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml".to_string()),
                ("/word/theme/theme1.xml".to_string(), "application/vnd.openxmlformats-officedocument.theme+xml".to_string()),
            ],
        }
    }

    /// Add image extension support
    pub fn add_image_extension(&mut self, ext: &str, content_type: &str) {
        // Check if extension already exists to avoid duplicates
        if !self.extensions.iter().any(|(e, _)| e == ext) {
            self.extensions
                .push((ext.to_string(), content_type.to_string()));
        }
    }

    /// Add numbering.xml
    pub fn add_numbering(&mut self) {
        self.overrides.push((
            "/word/numbering.xml".to_string(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
                .to_string(),
        ));
    }

    /// Add embedded font content type (.odttf)
    pub fn add_font_extension(&mut self) {
        if !self.extensions.iter().any(|(e, _)| e == "odttf") {
            self.extensions.push((
                "odttf".to_string(),
                "application/vnd.openxmlformats-officedocument.obfuscatedFont".to_string(),
            ));
        }
    }

    /// Add header
    pub fn add_header(&mut self, id: u32) {
        self.overrides.push((
            format!("/word/header{}.xml", id),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml".to_string(),
        ));
    }

    /// Add footer
    pub fn add_footer(&mut self, id: u32) {
        self.overrides.push((
            format!("/word/footer{}.xml", id),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml".to_string(),
        ));
    }

    /// Add footnotes.xml
    pub fn add_footnotes(&mut self) {
        self.overrides.push((
            "/word/footnotes.xml".to_string(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
                .to_string(),
        ));
    }

    /// Add endnotes.xml
    pub fn add_endnotes(&mut self) {
        self.overrides.push((
            "/word/endnotes.xml".to_string(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
                .to_string(),
        ));
    }

    /// Generate XML content
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        // XML declaration
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element
        let mut types = BytesStart::new("Types");
        types.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ));
        writer.write_event(Event::Start(types))?;

        // Default extensions
        for (ext, content_type) in &self.extensions {
            let mut default = BytesStart::new("Default");
            default.push_attribute(("Extension", ext.as_str()));
            default.push_attribute(("ContentType", content_type.as_str()));
            writer.write_event(Event::Empty(default))?;
        }

        // Overrides
        for (part_name, content_type) in &self.overrides {
            let mut override_elem = BytesStart::new("Override");
            override_elem.push_attribute(("PartName", part_name.as_str()));
            override_elem.push_attribute(("ContentType", content_type.as_str()));
            writer.write_event(Event::Empty(override_elem))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Types")))?;

        Ok(writer.into_inner().into_inner())
    }
}

impl Default for ContentTypes {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_content_types_basic() {
        let ct = ContentTypes::new();
        let xml = ct.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("Types"));
        assert!(xml_str.contains("Extension=\"rels\""));
        assert!(xml_str.contains("Extension=\"xml\""));
        assert!(xml_str.contains("PartName=\"/word/document.xml\""));
        // Check for docProps content types
        assert!(xml_str.contains("PartName=\"/docProps/core.xml\""));
        assert!(xml_str.contains("PartName=\"/docProps/app.xml\""));
        assert!(xml_str.contains("core-properties+xml"));
        assert!(xml_str.contains("extended-properties+xml"));
    }

    #[test]
    fn test_add_image_extension() {
        let mut ct = ContentTypes::new();
        ct.add_image_extension("png", "image/png");
        let xml = ct.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("Extension=\"png\""));
        assert!(xml_str.contains("image/png"));
    }

    #[test]
    fn test_add_numbering() {
        let mut ct = ContentTypes::new();
        ct.add_numbering();
        let xml = ct.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("PartName=\"/word/numbering.xml\""));
    }

    #[test]
    fn test_add_header_footer() {
        let mut ct = ContentTypes::new();
        ct.add_header(1);
        ct.add_footer(1);
        let xml = ct.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("PartName=\"/word/header1.xml\""));
        assert!(xml_str.contains("PartName=\"/word/footer1.xml\""));
    }

    #[test]
    fn test_add_footnotes() {
        let mut ct = ContentTypes::new();
        ct.add_footnotes();
        let xml = ct.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("PartName=\"/word/footnotes.xml\""));
        assert!(xml_str.contains("footnotes+xml"));
    }
}
