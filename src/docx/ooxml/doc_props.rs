//! Generate docProps/core.xml and docProps/app.xml for DOCX
//!
//! These files contain document metadata and are required for Word
//! to open the document without compatibility mode warnings.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;

/// Document metadata for core.xml
#[derive(Debug, Clone)]
pub struct CoreProperties {
    /// Document title
    pub title: Option<String>,
    /// Document subject
    pub subject: Option<String>,
    /// Document creator/author
    pub creator: Option<String>,
    /// Keywords/tags
    pub keywords: Option<String>,
    /// Description
    pub description: Option<String>,
    /// Last modified by
    pub last_modified_by: Option<String>,
    /// Revision number
    pub revision: Option<u32>,
    /// Creation date (ISO 8601 format)
    pub created: Option<String>,
    /// Last modified date (ISO 8601 format)
    pub modified: Option<String>,
}

impl Default for CoreProperties {
    fn default() -> Self {
        Self::new()
    }
}

impl CoreProperties {
    pub fn new() -> Self {
        // Get current time in ISO 8601 format
        let now = Self::current_iso_time();

        Self {
            title: None,
            subject: None,
            creator: Some("md2docx".to_string()),
            keywords: None,
            description: None,
            last_modified_by: Some("md2docx".to_string()),
            revision: Some(1),
            created: Some(now.clone()),
            modified: Some(now),
        }
    }

    /// Get current time in ISO 8601 format (W3CDTF)
    fn current_iso_time() -> String {
        // Use a fixed format that's compatible with Word
        // In production, this would use the system time
        // For now, use a reasonable default
        #[cfg(not(target_arch = "wasm32"))]
        {
            use std::time::SystemTime;
            let now = SystemTime::now()
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap_or_default();
            // Convert to ISO 8601 format manually (simplified)
            let secs = now.as_secs();
            // Calculate date components (simplified - assumes UTC)
            let days = secs / 86400;
            let years_since_1970 = days / 365;
            let year = 1970 + years_since_1970;
            let day_of_year = days % 365;
            let month = (day_of_year / 30).min(11) + 1;
            let day = (day_of_year % 30) + 1;
            let hour = (secs % 86400) / 3600;
            let minute = (secs % 3600) / 60;
            let second = secs % 60;
            format!(
                "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
                year, month, day, hour, minute, second
            )
        }
        #[cfg(target_arch = "wasm32")]
        {
            // Default timestamp for WASM
            "2025-01-01T00:00:00Z".to_string()
        }
    }

    /// Set document title
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Set document creator
    pub fn with_creator(mut self, creator: impl Into<String>) -> Self {
        self.creator = Some(creator.into());
        self
    }

    /// Generate core.xml content
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        // XML declaration
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with namespaces
        let mut root = BytesStart::new("cp:coreProperties");
        root.push_attribute((
            "xmlns:cp",
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        ));
        root.push_attribute(("xmlns:dc", "http://purl.org/dc/elements/1.1/"));
        root.push_attribute(("xmlns:dcterms", "http://purl.org/dc/terms/"));
        root.push_attribute(("xmlns:dcmitype", "http://purl.org/dc/dcmitype/"));
        root.push_attribute(("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"));
        writer.write_event(Event::Start(root))?;

        // Title
        if let Some(title) = &self.title {
            writer.write_event(Event::Start(BytesStart::new("dc:title")))?;
            writer.write_event(Event::Text(BytesText::new(title)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:title")))?;
        }

        // Subject
        if let Some(subject) = &self.subject {
            writer.write_event(Event::Start(BytesStart::new("dc:subject")))?;
            writer.write_event(Event::Text(BytesText::new(subject)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:subject")))?;
        }

        // Creator
        if let Some(creator) = &self.creator {
            writer.write_event(Event::Start(BytesStart::new("dc:creator")))?;
            writer.write_event(Event::Text(BytesText::new(creator)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:creator")))?;
        }

        // Keywords
        if let Some(keywords) = &self.keywords {
            writer.write_event(Event::Start(BytesStart::new("cp:keywords")))?;
            writer.write_event(Event::Text(BytesText::new(keywords)))?;
            writer.write_event(Event::End(BytesEnd::new("cp:keywords")))?;
        }

        // Description
        if let Some(description) = &self.description {
            writer.write_event(Event::Start(BytesStart::new("dc:description")))?;
            writer.write_event(Event::Text(BytesText::new(description)))?;
            writer.write_event(Event::End(BytesEnd::new("dc:description")))?;
        }

        // Last modified by
        if let Some(last_modified_by) = &self.last_modified_by {
            writer.write_event(Event::Start(BytesStart::new("cp:lastModifiedBy")))?;
            writer.write_event(Event::Text(BytesText::new(last_modified_by)))?;
            writer.write_event(Event::End(BytesEnd::new("cp:lastModifiedBy")))?;
        }

        // Revision
        if let Some(revision) = &self.revision {
            writer.write_event(Event::Start(BytesStart::new("cp:revision")))?;
            writer.write_event(Event::Text(BytesText::new(&revision.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("cp:revision")))?;
        }

        // Created date
        if let Some(created) = &self.created {
            let mut created_elem = BytesStart::new("dcterms:created");
            created_elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
            writer.write_event(Event::Start(created_elem))?;
            writer.write_event(Event::Text(BytesText::new(created)))?;
            writer.write_event(Event::End(BytesEnd::new("dcterms:created")))?;
        }

        // Modified date
        if let Some(modified) = &self.modified {
            let mut modified_elem = BytesStart::new("dcterms:modified");
            modified_elem.push_attribute(("xsi:type", "dcterms:W3CDTF"));
            writer.write_event(Event::Start(modified_elem))?;
            writer.write_event(Event::Text(BytesText::new(modified)))?;
            writer.write_event(Event::End(BytesEnd::new("dcterms:modified")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("cp:coreProperties")))?;

        Ok(writer.into_inner().into_inner())
    }
}

/// Application properties for app.xml
#[derive(Debug, Clone)]
pub struct AppProperties {
    /// Application name that created the document
    pub application: String,
    /// Application version
    pub app_version: String,
    /// Template used
    pub template: Option<String>,
    /// Total editing time in minutes
    pub total_time: Option<u32>,
    /// Page count
    pub pages: Option<u32>,
    /// Word count
    pub words: Option<u32>,
    /// Character count
    pub characters: Option<u32>,
    /// Character count with spaces
    pub characters_with_spaces: Option<u32>,
    /// Line count
    pub lines: Option<u32>,
    /// Paragraph count
    pub paragraphs: Option<u32>,
    /// Document security level
    pub doc_security: Option<u32>,
    /// Company name
    pub company: Option<String>,
}

impl Default for AppProperties {
    fn default() -> Self {
        Self::new()
    }
}

impl AppProperties {
    pub fn new() -> Self {
        Self {
            application: "md2docx".to_string(),
            app_version: env!("CARGO_PKG_VERSION").to_string(),
            template: Some("Normal.dotm".to_string()),
            total_time: Some(0),
            pages: None,
            words: None,
            characters: None,
            characters_with_spaces: None,
            lines: None,
            paragraphs: None,
            doc_security: Some(0),
            company: None,
        }
    }

    /// Generate app.xml content
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        // XML declaration
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // Root element with namespaces
        let mut root = BytesStart::new("Properties");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        ));
        root.push_attribute((
            "xmlns:vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        ));
        writer.write_event(Event::Start(root))?;

        // Template
        if let Some(template) = &self.template {
            writer.write_event(Event::Start(BytesStart::new("Template")))?;
            writer.write_event(Event::Text(BytesText::new(template)))?;
            writer.write_event(Event::End(BytesEnd::new("Template")))?;
        }

        // TotalTime
        if let Some(total_time) = &self.total_time {
            writer.write_event(Event::Start(BytesStart::new("TotalTime")))?;
            writer.write_event(Event::Text(BytesText::new(&total_time.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("TotalTime")))?;
        }

        // Pages
        if let Some(pages) = &self.pages {
            writer.write_event(Event::Start(BytesStart::new("Pages")))?;
            writer.write_event(Event::Text(BytesText::new(&pages.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("Pages")))?;
        }

        // Words
        if let Some(words) = &self.words {
            writer.write_event(Event::Start(BytesStart::new("Words")))?;
            writer.write_event(Event::Text(BytesText::new(&words.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("Words")))?;
        }

        // Characters
        if let Some(characters) = &self.characters {
            writer.write_event(Event::Start(BytesStart::new("Characters")))?;
            writer.write_event(Event::Text(BytesText::new(&characters.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("Characters")))?;
        }

        // Application
        writer.write_event(Event::Start(BytesStart::new("Application")))?;
        writer.write_event(Event::Text(BytesText::new(&self.application)))?;
        writer.write_event(Event::End(BytesEnd::new("Application")))?;

        // DocSecurity
        if let Some(doc_security) = &self.doc_security {
            writer.write_event(Event::Start(BytesStart::new("DocSecurity")))?;
            writer.write_event(Event::Text(BytesText::new(&doc_security.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("DocSecurity")))?;
        }

        // Lines
        if let Some(lines) = &self.lines {
            writer.write_event(Event::Start(BytesStart::new("Lines")))?;
            writer.write_event(Event::Text(BytesText::new(&lines.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("Lines")))?;
        }

        // Paragraphs
        if let Some(paragraphs) = &self.paragraphs {
            writer.write_event(Event::Start(BytesStart::new("Paragraphs")))?;
            writer.write_event(Event::Text(BytesText::new(&paragraphs.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("Paragraphs")))?;
        }

        // ScaleCrop (required)
        writer.write_event(Event::Start(BytesStart::new("ScaleCrop")))?;
        writer.write_event(Event::Text(BytesText::new("false")))?;
        writer.write_event(Event::End(BytesEnd::new("ScaleCrop")))?;

        // Company
        if let Some(company) = &self.company {
            writer.write_event(Event::Start(BytesStart::new("Company")))?;
            writer.write_event(Event::Text(BytesText::new(company)))?;
            writer.write_event(Event::End(BytesEnd::new("Company")))?;
        } else {
            // Write empty Company element (Word expects this)
            writer.write_event(Event::Start(BytesStart::new("Company")))?;
            writer.write_event(Event::End(BytesEnd::new("Company")))?;
        }

        // LinksUpToDate (required)
        writer.write_event(Event::Start(BytesStart::new("LinksUpToDate")))?;
        writer.write_event(Event::Text(BytesText::new("false")))?;
        writer.write_event(Event::End(BytesEnd::new("LinksUpToDate")))?;

        // CharactersWithSpaces
        if let Some(chars_with_spaces) = &self.characters_with_spaces {
            writer.write_event(Event::Start(BytesStart::new("CharactersWithSpaces")))?;
            writer.write_event(Event::Text(BytesText::new(&chars_with_spaces.to_string())))?;
            writer.write_event(Event::End(BytesEnd::new("CharactersWithSpaces")))?;
        }

        // SharedDoc (required)
        writer.write_event(Event::Start(BytesStart::new("SharedDoc")))?;
        writer.write_event(Event::Text(BytesText::new("false")))?;
        writer.write_event(Event::End(BytesEnd::new("SharedDoc")))?;

        // HyperlinksChanged (required)
        writer.write_event(Event::Start(BytesStart::new("HyperlinksChanged")))?;
        writer.write_event(Event::Text(BytesText::new("false")))?;
        writer.write_event(Event::End(BytesEnd::new("HyperlinksChanged")))?;

        // AppVersion - format as major.minor (e.g., "1.0000")
        writer.write_event(Event::Start(BytesStart::new("AppVersion")))?;
        // Parse version and format as Word expects
        let version_parts: Vec<&str> = self.app_version.split('.').collect();
        let major = version_parts.first().unwrap_or(&"1");
        let minor = version_parts.get(1).unwrap_or(&"0");
        let formatted_version = format!("{}.{:0<4}", major, minor);
        writer.write_event(Event::Text(BytesText::new(&formatted_version)))?;
        writer.write_event(Event::End(BytesEnd::new("AppVersion")))?;

        writer.write_event(Event::End(BytesEnd::new("Properties")))?;

        Ok(writer.into_inner().into_inner())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_core_properties_default() {
        let core = CoreProperties::new();
        let xml = core.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("cp:coreProperties"));
        assert!(xml_str.contains("dc:creator"));
        assert!(xml_str.contains("md2docx"));
        assert!(xml_str.contains("dcterms:created"));
        assert!(xml_str.contains("dcterms:modified"));
    }

    #[test]
    fn test_core_properties_with_title() {
        let core = CoreProperties::new().with_title("My Document");
        let xml = core.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("dc:title"));
        assert!(xml_str.contains("My Document"));
    }

    #[test]
    fn test_app_properties_default() {
        let app = AppProperties::new();
        let xml = app.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("Properties"));
        assert!(xml_str.contains("Application"));
        assert!(xml_str.contains("md2docx"));
        assert!(xml_str.contains("AppVersion"));
        assert!(xml_str.contains("ScaleCrop"));
        assert!(xml_str.contains("LinksUpToDate"));
    }

    #[test]
    fn test_app_properties_has_required_elements() {
        let app = AppProperties::new();
        let xml = app.to_xml().unwrap();
        let xml_str = String::from_utf8(xml).unwrap();

        // These are required by Word
        assert!(xml_str.contains("<ScaleCrop>false</ScaleCrop>"));
        assert!(xml_str.contains("<LinksUpToDate>false</LinksUpToDate>"));
        assert!(xml_str.contains("<SharedDoc>false</SharedDoc>"));
        assert!(xml_str.contains("<HyperlinksChanged>false</HyperlinksChanged>"));
    }
}
