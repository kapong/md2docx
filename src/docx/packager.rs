//! ZIP packager for DOCX

use std::io::{Seek, Write};
use zip::write::{FileOptions, ZipWriter};

use crate::docx::ooxml::{
    generate_font_table_xml, generate_settings_xml, generate_theme_xml, generate_web_settings_xml,
    AppProperties, ContentTypes, CoreProperties, DocumentXml, Language, Relationships,
    StylesDocument,
};
use crate::error::Result;

/// DOCX Packager
///
/// Assembles all OOXML components into a valid DOCX (ZIP) file.
pub(crate) struct Packager<W: Write + Seek> {
    writer: ZipWriter<W>,
    added_files: std::collections::HashSet<String>,
}


/// Custom document properties for packaging
pub(crate) struct DocProps<'a> {
    pub core: &'a CoreProperties,
    pub app: &'a AppProperties,
}

/// Relationships context for packaging
pub(crate) struct RelContext<'a> {
    pub(crate) root: &'a Relationships,
    pub(crate) doc: &'a Relationships,
}

impl<W: Write + Seek> Packager<W> {
    /// Create a new packager with the given writer
    pub fn new(writer: W) -> Self {
        Self {
            writer: ZipWriter::new(writer),
            added_files: std::collections::HashSet::new(),
        }
    }

    /// Get file options for writing
    fn get_file_options() -> FileOptions<'static, ()> {
        FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o644)
    }

    /// Package all DOCX components into the ZIP archive
    ///
    /// This writes all required OOXML files to the correct paths:
    /// - `[Content_Types].xml`
    /// - `_rels/.rels`
    /// - `docProps/core.xml`
    /// - `docProps/app.xml`
    /// - `word/document.xml`
    /// - `word/styles.xml`
    /// - `word/settings.xml`
    /// - `word/fontTable.xml`
    /// - `word/_rels/document.xml.rels`
    pub(crate) fn package(
        &mut self,
        document: &DocumentXml,
        styles: &StylesDocument,
        content_types: &ContentTypes,
        rels: &Relationships,     // _rels/.rels
        doc_rels: &Relationships, // word/_rels/document.xml.rels
        lang: Language,
    ) -> Result<()> {
        // Use default document properties
        let core_props = CoreProperties::new();
        let app_props = AppProperties::new();
        self.package_with_props(
            document,
            styles,
            content_types,
            &RelContext {
                root: rels,
                doc: doc_rels,
            },
            lang,
            &DocProps {
                core: &core_props,
                app: &app_props,
            },
        )
    }

    /// Package all DOCX components with custom document properties
    pub(crate) fn package_with_props(
        &mut self,
        document: &DocumentXml,
        styles: &StylesDocument,
        content_types: &ContentTypes,
        rels: &RelContext,
        lang: Language,
        props: &DocProps,
    ) -> Result<()> {
        // 1. [Content_Types].xml - Defines content types for all parts
        self.write_file("[Content_Types].xml", &content_types.to_xml()?)?;

        // 2. _rels/.rels - Root relationships (points to document.xml, docProps)
        self.write_file("_rels/.rels", &rels.root.to_xml()?)?;

        // 3. docProps/core.xml - Core document properties (author, title, dates)
        self.write_file("docProps/core.xml", &props.core.to_xml()?)?;

        // 4. docProps/app.xml - Application properties (creator app, version)
        self.write_file("docProps/app.xml", &props.app.to_xml()?)?;

        // 5. word/document.xml - Main document content
        self.write_file("word/document.xml", &document.to_xml()?)?;

        // 6. word/styles.xml - Style definitions
        self.write_file("word/styles.xml", &styles.to_xml()?)?;

        // 7. word/settings.xml - Document settings
        self.write_file("word/settings.xml", &generate_settings_xml()?)?;

        // 8. word/fontTable.xml - Font table
        self.write_file("word/fontTable.xml", &generate_font_table_xml(lang)?)?;

        // 9. word/webSettings.xml - Web settings (required for Word compatibility)
        self.write_file("word/webSettings.xml", &generate_web_settings_xml()?)?;

        // 10. word/theme/theme1.xml - Theme (required for Word compatibility)
        self.write_file("word/theme/theme1.xml", &generate_theme_xml()?)?;

        // 11. word/_rels/document.xml.rels - Document relationships
        self.write_file("word/_rels/document.xml.rels", &rels.doc.to_xml()?)?;

        Ok(())
    }

    /// Write a file to the ZIP archive
    fn write_file(&mut self, path: &str, content: &[u8]) -> Result<()> {
        if self.added_files.contains(path) {
            return Ok(());
        }
        self.writer.start_file(path, Self::get_file_options())?;
        self.writer.write_all(content)?;
        self.added_files.insert(path.to_string());
        Ok(())
    }

    /// Add an image file to the archive
    ///
    /// Images are stored in `word/media/` directory.
    pub fn add_image(&mut self, filename: &str, content: &[u8]) -> Result<()> {
        let path = format!("word/media/{}", filename);
        self.write_file(&path, content)?;
        Ok(())
    }

    /// Add a header file to the archive
    pub fn add_header(&mut self, header_num: u32, content: &[u8]) -> Result<()> {
        let path = format!("word/header{}.xml", header_num);
        self.write_file(&path, content)?;
        Ok(())
    }

    /// Add a footer file to the archive
    pub fn add_footer(&mut self, footer_num: u32, content: &[u8]) -> Result<()> {
        let path = format!("word/footer{}.xml", footer_num);
        self.write_file(&path, content)?;
        Ok(())
    }

    /// Add a header relationships file to the archive
    pub fn add_header_rels(&mut self, header_num: u32, content: &[u8]) -> Result<()> {
        let path = format!("word/_rels/header{}.xml.rels", header_num);
        self.write_file(&path, content)?;
        Ok(())
    }

    /// Add a footer relationships file to the archive
    pub fn add_footer_rels(&mut self, footer_num: u32, content: &[u8]) -> Result<()> {
        let path = format!("word/_rels/footer{}.xml.rels", footer_num);
        self.write_file(&path, content)?;
        Ok(())
    }

    /// Add a numbering file to the archive
    pub fn add_numbering(&mut self, content: &[u8]) -> Result<()> {
        self.write_file("word/numbering.xml", content)?;
        Ok(())
    }

    /// Add a footnotes file to the archive
    pub fn add_footnotes(&mut self, content: &[u8]) -> Result<()> {
        self.write_file("word/footnotes.xml", content)?;
        Ok(())
    }

    /// Add an endnotes file to the archive
    pub fn add_endnotes(&mut self, content: &[u8]) -> Result<()> {
        self.write_file("word/endnotes.xml", content)?;
        Ok(())
    }

    /// Finish writing the ZIP archive
    ///
    /// This must be called after all files have been added.
    /// Consumes the packager and returns the underlying writer.
    pub fn finish(self) -> Result<W> {
        let writer = self.writer.finish()?;
        Ok(writer)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_packager_basic() {
        // Create test components
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        // Create packager with in-memory buffer
        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package the components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Finish and get the ZIP data
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        // Verify we got some data
        assert!(!zip_data.is_empty());

        // Verify it's a valid ZIP by checking for ZIP magic bytes
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_image() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add an image
        let image_data = b"fake image data";
        packager.add_image("test.png", image_data).unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_thai() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::Thai, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::Thai,
            )
            .unwrap();

        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_header_footer() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add header and footer
        let header_data = b"<w:hdr><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>";
        let footer_data = b"<w:ftr><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>";

        packager.add_header(1, header_data).unwrap();
        packager.add_footer(1, footer_data).unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_numbering() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add numbering
        let numbering_data = b"<w:numbering><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>";
        packager.add_numbering(numbering_data).unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_document_content() {
        use crate::docx::ooxml::Paragraph;

        // Create a document with actual content
        let mut document = DocumentXml::new();
        document.add_paragraph(Paragraph::with_style("Heading1").add_text("Test Document"));
        document
            .add_paragraph(Paragraph::with_style("Normal").add_text("This is a test paragraph."));

        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_multiple_images() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add multiple images
        packager.add_image("image1.png", b"image1 data").unwrap();
        packager.add_image("image2.jpg", b"image2 data").unwrap();
        packager.add_image("image3.png", b"image3 data").unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_footnotes() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add footnotes
        let footnotes_data = b"<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id=\"1\"><w:p><w:r><w:t>This is a footnote</w:t></w:r></w:p></w:footnote></w:footnotes>";
        packager.add_footnotes(footnotes_data).unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_header_rels() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add header rels
        let rels_data = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;
        packager.add_header_rels(1, rels_data).unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }

    #[test]
    fn test_packager_with_footer_rels() {
        let document = DocumentXml::new();
        let styles = StylesDocument::new(Language::English, None);
        let content_types = ContentTypes::new();
        let rels = Relationships::root_rels();
        let doc_rels = Relationships::document_rels();

        let buffer = Cursor::new(Vec::new());
        let mut packager = Packager::new(buffer);

        // Package components
        packager
            .package(
                &document,
                &styles,
                &content_types,
                &rels,
                &doc_rels,
                Language::English,
            )
            .unwrap();

        // Add footer rels
        let rels_data = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
</Relationships>"#;
        packager.add_footer_rels(1, rels_data).unwrap();

        // Finish
        let buffer = packager.finish().unwrap();
        let zip_data = buffer.into_inner();

        assert!(!zip_data.is_empty());
        assert_eq!(&zip_data[0..4], b"PK\x03\x04");
    }
}
