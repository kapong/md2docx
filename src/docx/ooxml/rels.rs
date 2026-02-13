//! Generate relationship files for DOCX

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;

/// Relationship entry
#[derive(Clone)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>, // For external links
}

/// Relationships container
pub(crate) struct Relationships {
    rels: Vec<Relationship>,
    next_id: usize,
}

impl Relationships {
    pub fn new() -> Self {
        Self {
            rels: Vec::new(),
            next_id: 1,
        }
    }

    /// Create root .rels file (points to document.xml, core.xml, app.xml)
    pub fn root_rels() -> Self {
        let mut rels = Self::new();
        // rId1: Main document
        rels.add(Relationship {
            id: "rId1".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    .to_string(),
            target: "word/document.xml".to_string(),
            target_mode: None,
        });
        // rId2: Core properties (author, title, dates)
        rels.add(Relationship {
            id: "rId2".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                    .to_string(),
            target: "docProps/core.xml".to_string(),
            target_mode: None,
        });
        // rId3: Extended/app properties (application, version)
        rels.add(Relationship {
            id: "rId3".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                    .to_string(),
            target: "docProps/app.xml".to_string(),
            target_mode: None,
        });
        rels
    }

    /// Create document.xml.rels (styles, settings, webSettings, theme, etc.)
    pub fn document_rels() -> Self {
        let mut rels = Self::new();
        rels.add(Relationship {
            id: "rId1".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                .to_string(),
            target: "styles.xml".to_string(),
            target_mode: None,
        });
        rels.add(Relationship {
            id: "rId2".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
                    .to_string(),
            target: "settings.xml".to_string(),
            target_mode: None,
        });
        rels.add(Relationship {
            id: "rId3".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
                    .to_string(),
            target: "fontTable.xml".to_string(),
            target_mode: None,
        });
        // rId4: webSettings (required for Word compatibility)
        rels.add(Relationship {
            id: "rId4".to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
                    .to_string(),
            target: "webSettings.xml".to_string(),
            target_mode: None,
        });
        // rId5: theme (required for Word compatibility)
        rels.add(Relationship {
            id: "rId5".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                .to_string(),
            target: "theme/theme1.xml".to_string(),
            target_mode: None,
        });
        rels
    }

    /// Add a relationship
    pub fn add(&mut self, rel: Relationship) {
        // Update next_id based on added ID to prevent collisions
        if rel.id.starts_with("rId") {
            if let Ok(num) = rel.id[3..].parse::<usize>() {
                if num >= self.next_id {
                    self.next_id = num + 1;
                }
            }
        }
        self.rels.push(rel);
    }

    /// Add and return the relationship ID
    pub fn add_and_get_id(&mut self, rel_type: &str, target: &str) -> String {
        let id = format!("rId{}", self.next_id);
        self.add(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
        id
    }

    /// Add external hyperlink
    #[allow(dead_code)]
    pub fn add_hyperlink(&mut self, url: &str) -> String {
        let id = format!("rId{}", self.next_id);
        self.add(Relationship {
            id: id.clone(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                    .to_string(),
            target: url.to_string(),
            target_mode: Some("External".to_string()),
        });
        id
    }

    /// Add external hyperlink with custom ID
    pub fn add_hyperlink_with_id(&mut self, id: &str, url: &str) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                    .to_string(),
            target: url.to_string(),
            target_mode: Some("External".to_string()),
        });
    }

    /// Add image with auto-generated ID
    #[allow(dead_code)]
    pub fn add_image(&mut self, filename: &str) -> String {
        let id = format!("rId{}", self.next_id);
        self.add(Relationship {
            id: id.clone(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                .to_string(),
            target: format!("media/{}", filename),
            target_mode: None,
        });
        id
    }

    /// Add image with specific ID (needed when syncing with ImageContext)
    pub fn add_image_with_id(&mut self, id: &str, filename: &str) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                .to_string(),
            target: format!("media/{}", filename),
            target_mode: None,
        });
    }

    /// Add header with auto-generated ID
    pub fn add_header(&mut self, header_num: u32) -> String {
        let id = format!("rId{}", self.next_id);
        self.add_header_with_id(&id, header_num);
        id
    }

    /// Add header with specific ID
    pub fn add_header_with_id(&mut self, id: &str, header_num: u32) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
                .to_string(),
            target: format!("header{}.xml", header_num),
            target_mode: None,
        });
    }

    /// Add footer with auto-generated ID
    pub fn add_footer(&mut self, footer_num: u32) -> String {
        let id = format!("rId{}", self.next_id);
        self.add_footer_with_id(&id, footer_num);
        id
    }

    /// Add footer with specific ID
    pub fn add_footer_with_id(&mut self, id: &str, footer_num: u32) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
                .to_string(),
            target: format!("footer{}.xml", footer_num),
            target_mode: None,
        });
    }

    /// Add numbering with auto-generated ID
    pub fn add_numbering(&mut self) -> String {
        self.add_and_get_id(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
            "numbering.xml",
        )
    }

    /// Add numbering with specific ID
    pub fn add_numbering_with_id(&mut self, id: &str) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
                    .to_string(),
            target: "numbering.xml".to_string(),
            target_mode: None,
        });
    }

    /// Add footnotes
    pub fn add_footnotes(&mut self) -> String {
        self.add_and_get_id(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
            "footnotes.xml",
        )
    }

    /// Add footnotes with specific ID
    pub fn add_footnotes_with_id(&mut self, id: &str) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
                    .to_string(),
            target: "footnotes.xml".to_string(),
            target_mode: None,
        });
    }

    /// Add endnotes
    pub fn add_endnotes(&mut self) -> String {
        self.add_and_get_id(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
            "endnotes.xml",
        )
    }

    /// Add endnotes with specific ID
    pub fn add_endnotes_with_id(&mut self, id: &str) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
                    .to_string(),
            target: "endnotes.xml".to_string(),
            target_mode: None,
        });
    }

    /// Add embedded font relationship (for fontTable.xml.rels)
    pub fn add_font_with_id(&mut self, id: &str, filename: &str) {
        self.add(Relationship {
            id: id.to_string(),
            rel_type:
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
                    .to_string(),
            target: format!("fonts/{}", filename),
            target_mode: None,
        });
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
        let mut relationships = BytesStart::new("Relationships");
        relationships.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(relationships))?;

        // Each relationship
        for rel in &self.rels {
            let mut rel_elem = BytesStart::new("Relationship");
            rel_elem.push_attribute(("Id", rel.id.as_str()));
            rel_elem.push_attribute(("Type", rel.rel_type.as_str()));
            rel_elem.push_attribute(("Target", rel.target.as_str()));
            if let Some(mode) = &rel.target_mode {
                rel_elem.push_attribute(("TargetMode", mode.as_str()));
            }
            writer.write_event(Event::Empty(rel_elem))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;

        Ok(writer.into_inner().into_inner())
    }
}

impl Default for Relationships {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_relationships_basic() {
        let rels = Relationships::new();
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("Relationships"));
        assert!(xml_str.contains("xmlns"));
    }

    #[test]
    fn test_root_rels() {
        let rels = Relationships::root_rels();
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("Id=\"rId1\""));
        assert!(xml_str.contains("word/document.xml"));
        assert!(xml_str.contains("officeDocument"));
        // Check for docProps relationships
        assert!(xml_str.contains("Id=\"rId2\""));
        assert!(xml_str.contains("docProps/core.xml"));
        assert!(xml_str.contains("core-properties"));
        assert!(xml_str.contains("Id=\"rId3\""));
        assert!(xml_str.contains("docProps/app.xml"));
        assert!(xml_str.contains("extended-properties"));
    }

    #[test]
    fn test_document_rels() {
        let rels = Relationships::document_rels();
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("styles.xml"));
        assert!(xml_str.contains("settings.xml"));
        assert!(xml_str.contains("fontTable.xml"));
        // Check for webSettings and theme
        assert!(xml_str.contains("webSettings.xml"));
        assert!(xml_str.contains("theme/theme1.xml"));
    }

    #[test]
    fn test_add_hyperlink() {
        let mut rels = Relationships::new();
        let id = rels.add_hyperlink("https://example.com");
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert_eq!(id, "rId1");
        assert!(xml_str.contains("hyperlink"));
        assert!(xml_str.contains("TargetMode=\"External\""));
        assert!(xml_str.contains("https://example.com"));
    }

    #[test]
    fn test_add_image() {
        let mut rels = Relationships::new();
        let id = rels.add_image("image1.png");
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert_eq!(id, "rId1");
        assert!(xml_str.contains("image"));
        assert!(xml_str.contains("media/image1.png"));
    }

    #[test]
    fn test_add_header_footer() {
        let mut rels = Relationships::new();
        let header_id = rels.add_header(1);
        let footer_id = rels.add_footer(1);
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert_eq!(header_id, "rId1");
        assert_eq!(footer_id, "rId2");
        assert!(xml_str.contains("header1.xml"));
        assert!(xml_str.contains("footer1.xml"));
    }

    #[test]
    fn test_add_numbering() {
        let mut rels = Relationships::new();
        let id = rels.add_numbering();
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert_eq!(id, "rId1");
        assert!(xml_str.contains("numbering.xml"));
    }

    #[test]
    fn test_add_footnotes() {
        let mut rels = Relationships::new();
        let id = rels.add_footnotes();
        let xml = rels.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert_eq!(id, "rId1");
        assert!(xml_str.contains("footnotes.xml"));
        assert!(xml_str.contains("relationships/footnotes"));
    }
}
