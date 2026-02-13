//! Header/Footer template extraction from DOCX files

use crate::error::{Error, Result};
use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

/// Represents an extracted header/footer template
#[derive(Debug, Clone, Default)]
pub struct HeaderFooterTemplate {
    /// Default header content (for pages after first)
    pub default_header: Option<HeaderFooterContent>,
    /// Default footer content (for pages after first)
    pub default_footer: Option<HeaderFooterContent>,
    /// First page header content (when different_first_page is true)
    pub first_page_header: Option<HeaderFooterContent>,
    /// First page footer content (when different_first_page is true)
    pub first_page_footer: Option<HeaderFooterContent>,
    /// Whether first page is different (w:titlePg flag was set in document)
    pub different_first_page: bool,
    /// Media files referenced by headers/footers (images, etc.)
    pub media: Vec<MediaFile>,
    /// Tab stops from the template's Header style (position, alignment)
    pub header_style_tabs: Vec<(u32, String)>,
    /// Tab stops from the template's Footer style (position, alignment)
    pub footer_style_tabs: Vec<(u32, String)>,
}

/// Content of a single header or footer
#[derive(Debug, Clone)]
pub struct HeaderFooterContent {
    /// Raw XML content (everything inside <w:hdr> or <w:ftr>, including the root element)
    pub raw_xml: String,
    /// Detected placeholders in the content (e.g., ["title", "page", "chapter"])
    pub placeholders: Vec<String>,
    /// Relationship ID mappings from this header/footer's rels file (rId -> target path)
    pub rel_id_map: HashMap<String, String>,
}

/// Media file extracted from the template
#[derive(Debug, Clone)]
pub struct MediaFile {
    /// Filename (e.g., "image1.png")
    pub filename: String,
    /// File content bytes
    pub data: Vec<u8>,
    /// Content type (e.g., "image/png")
    pub content_type: String,
}

impl HeaderFooterTemplate {
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if template has any content
    pub fn is_empty(&self) -> bool {
        self.default_header.is_none()
            && self.default_footer.is_none()
            && self.first_page_header.is_none()
            && self.first_page_footer.is_none()
    }

    /// Get all unique placeholders from all headers/footers
    pub fn all_placeholders(&self) -> Vec<String> {
        let mut keys = std::collections::HashSet::new();

        if let Some(ref h) = self.default_header {
            for k in &h.placeholders {
                keys.insert(k.clone());
            }
        }
        if let Some(ref f) = self.default_footer {
            for k in &f.placeholders {
                keys.insert(k.clone());
            }
        }
        if let Some(ref h) = self.first_page_header {
            for k in &h.placeholders {
                keys.insert(k.clone());
            }
        }
        if let Some(ref f) = self.first_page_footer {
            for k in &f.placeholders {
                keys.insert(k.clone());
            }
        }

        keys.into_iter().collect()
    }
}

/// Extract header/footer template from a DOCX file
pub fn extract(path: &Path) -> Result<HeaderFooterTemplate> {
    if !path.exists() {
        return Err(Error::Template(format!(
            "Header/footer template file not found: {}",
            path.display()
        )));
    }

    let file = std::fs::File::open(path)?;
    let mut archive = zip::ZipArchive::new(file)?;

    // 1. Read document.xml.rels to find header/footer files
    let doc_rels = read_archive_file(&mut archive, "word/_rels/document.xml.rels")?;

    // 2. Parse relationships to find header/footer file mappings
    let _header_files = find_header_footer_files(&doc_rels, "header");
    let _footer_files = find_header_footer_files(&doc_rels, "footer");

    // 3. Read document.xml to check for w:titlePg (different first page) and header/footer references
    let document_xml = read_archive_file(&mut archive, "word/document.xml")?;
    let different_first_page =
        document_xml.contains("<w:titlePg") || document_xml.contains("<w:titlePg/>");

    // 4. Determine which files are default vs first page
    // Note: headerReference/footerReference elements are in document.xml, but we need
    // doc_rels to map rId to actual file paths
    let (default_header_file, first_header_file) =
        categorize_header_footer_files(&document_xml, &doc_rels, "header");
    let (default_footer_file, first_footer_file) =
        categorize_header_footer_files(&document_xml, &doc_rels, "footer");

    // 5. Extract each header/footer content
    let default_header = extract_header_footer_content(&mut archive, &default_header_file, "word")?;
    let first_page_header = if different_first_page {
        extract_header_footer_content(&mut archive, &first_header_file, "word")?
    } else {
        None
    };

    let default_footer = extract_header_footer_content(&mut archive, &default_footer_file, "word")?;
    let first_page_footer = if different_first_page {
        extract_header_footer_content(&mut archive, &first_footer_file, "word")?
    } else {
        None
    };

    // 6. Collect all media files referenced by headers/footers
    let mut media = Vec::new();
    collect_media_files(&mut archive, &default_header, &mut media)?;
    collect_media_files(&mut archive, &first_page_header, &mut media)?;
    collect_media_files(&mut archive, &default_footer, &mut media)?;
    collect_media_files(&mut archive, &first_page_footer, &mut media)?;

    // 7. Extract Header/Footer style tab stops from styles.xml
    let (header_style_tabs, footer_style_tabs) = extract_style_tabs(&mut archive)?;

    Ok(HeaderFooterTemplate {
        default_header,
        default_footer,
        first_page_header,
        first_page_footer,
        different_first_page,
        media,
        header_style_tabs,
        footer_style_tabs,
    })
}

/// Read a file from the ZIP archive as a string
fn read_archive_file<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<String> {
    let mut file = archive
        .by_name(path)
        .map_err(|e| Error::Zip(format!("Failed to read {}: {}", path, e)))?;
    let mut content = String::new();
    file.read_to_string(&mut content).map_err(Error::Io)?;
    Ok(content)
}

/// Find all header or footer files referenced in relationships XML
fn find_header_footer_files(rels_xml: &str, hf_type: &str) -> Vec<String> {
    let mut files = Vec::new();

    // Look for Relationship elements with Type containing "header" or "footer"
    let relationship_regex = regex::Regex::new(
        r#"<Relationship[^>]*Type="[^"]*/(header|footer)"[^>]*Target="([^"]*)"[^>]*/>"#,
    )
    .expect("relationship_regex should be valid");

    for cap in relationship_regex.captures_iter(rels_xml) {
        let type_match = cap
            .get(1)
            .expect("relationship_regex should have capture group 1")
            .as_str();
        let target = cap
            .get(2)
            .expect("relationship_regex should have capture group 2")
            .as_str();

        if type_match == hf_type {
            // Target might be like "header1.xml" or "header2.xml"
            files.push(target.to_string());
        }
    }

    files
}

/// Categorize header/footer files into default and first page
///
/// Returns (default_file, first_page_file)
fn categorize_header_footer_files(
    document_xml: &str,
    rels_xml: &str,
    hf_type: &str,
) -> (Option<String>, Option<String>) {
    let mut default_file = None;
    let mut first_page_file = None;

    // Look for headerReference or footerReference elements in document.xml
    // These have w:type="default" or w:type="first"
    // Pattern: <w:headerReference w:type="default" r:id="rId8"/>
    let reference_regex = regex::Regex::new(
        r#"<w:(header|footer)Reference[^>]*w:type="(default|first|even)"[^>]*r:id="([^"]*)"[^>]*/>"#,
    )
    .expect("reference_regex should be valid");

    for cap in reference_regex.captures_iter(document_xml) {
        let ref_type = cap
            .get(1)
            .expect("reference_regex should have capture group 1")
            .as_str();
        let ref_subtype = cap
            .get(2)
            .expect("reference_regex should have capture group 2")
            .as_str();
        let r_id = cap
            .get(3)
            .expect("reference_regex should have capture group 3")
            .as_str();

        if ref_type != hf_type {
            continue;
        }

        // Find the target for this rId in the rels file
        if let Some(target) = find_target_by_rid(rels_xml, r_id) {
            if ref_subtype == "default" {
                default_file = Some(target);
            } else if ref_subtype == "first" {
                first_page_file = Some(target);
            }
            // "even" type is for even pages - we skip it for now
        }
    }

    // If we couldn't find explicit type attributes, use the first file as default
    if default_file.is_none() {
        let files = find_header_footer_files(rels_xml, hf_type);
        if !files.is_empty() {
            default_file = Some(files[0].clone());
            if files.len() > 1 {
                first_page_file = Some(files[1].clone());
            }
        }
    }

    (default_file, first_page_file)
}

/// Find the Target for a given rId in relationships XML
fn find_target_by_rid(rels_xml: &str, r_id: &str) -> Option<String> {
    let regex = regex::Regex::new(&format!(
        r#"<Relationship[^>]*Id="{}"[^>]*Target="([^"]*)"[^>]*/>"#,
        regex::escape(r_id)
    ))
    .expect("find_target_by_rid regex should be valid");

    regex.captures(rels_xml).map(|cap| {
        cap.get(1)
            .expect("find_target_by_rid regex should have capture group 1")
            .as_str()
            .to_string()
    })
}

/// Extract header/footer content from a file in the archive
fn extract_header_footer_content<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    filename: &Option<String>,
    base_path: &str,
) -> Result<Option<HeaderFooterContent>> {
    let filename = match filename {
        Some(f) => f,
        None => return Ok(None),
    };

    // Construct full path (e.g., "word/header1.xml")
    let full_path = if filename.starts_with(base_path) {
        filename.clone()
    } else {
        format!("{}/{}", base_path, filename)
    };

    // Read the header/footer XML
    let xml = read_archive_file(archive, &full_path)?;

    // Extract placeholders from the XML
    let placeholders = extract_placeholders_from_xml(&xml);

    // Read the rels file for this header/footer if it exists
    let rel_id_map = extract_rel_id_map(archive, &full_path)?;

    Ok(Some(HeaderFooterContent {
        raw_xml: xml,
        placeholders,
        rel_id_map,
    }))
}

/// Extract placeholders from XML content
///
/// Looks for {{placeholder}} patterns within <w:t> elements
pub fn extract_placeholders_from_xml(xml: &str) -> Vec<String> {
    let mut placeholders = Vec::new();
    let placeholder_regex =
        regex::Regex::new(r"\{\{(\w+)\}\}").expect("placeholder_regex should be valid");

    for cap in placeholder_regex.captures_iter(xml) {
        let key = cap
            .get(1)
            .expect("placeholder_regex should have capture group 1")
            .as_str()
            .to_string();
        if !placeholders.contains(&key) {
            placeholders.push(key);
        }
    }

    placeholders
}

/// Extract relationship ID mappings from a header/footer's rels file
fn extract_rel_id_map<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    header_footer_path: &str,
) -> Result<HashMap<String, String>> {
    let mut rel_id_map = HashMap::new();

    // Construct rels file path (e.g., "word/_rels/header1.xml.rels")
    let filename = header_footer_path
        .strip_prefix("word/")
        .unwrap_or(header_footer_path);
    let rels_path = format!("word/_rels/{}.rels", filename);

    // Try to read the rels file
    let rels_xml = match read_archive_file(archive, &rels_path) {
        Ok(xml) => xml,
        Err(_) => return Ok(rel_id_map), // No rels file is OK
    };

    // Parse relationships
    let relationship_regex =
        regex::Regex::new(r#"<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*/>"#)
            .expect("relationship_regex should be valid");

    for cap in relationship_regex.captures_iter(&rels_xml) {
        let r_id = cap
            .get(1)
            .expect("relationship_regex should have capture group 1")
            .as_str()
            .to_string();
        let target = cap
            .get(2)
            .expect("relationship_regex should have capture group 2")
            .as_str()
            .to_string();
        rel_id_map.insert(r_id, target);
    }

    Ok(rel_id_map)
}

/// Collect media files referenced by header/footer content
fn collect_media_files<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    content: &Option<HeaderFooterContent>,
    media: &mut Vec<MediaFile>,
) -> Result<()> {
    let content = match content {
        Some(c) => c,
        None => return Ok(()),
    };

    // Find all r:embed references in the XML
    let embed_regex =
        regex::Regex::new(r#"r:embed="([^"]*)""#).expect("embed_regex pattern should be valid");

    for cap in embed_regex.captures_iter(&content.raw_xml) {
        let r_id = cap
            .get(1)
            .expect("embed_regex should have capture group 1")
            .as_str();

        // Look up the target in the rel_id_map
        if let Some(target) = content.rel_id_map.get(r_id) {
            // Construct full path to media file (e.g., "word/media/image1.png")
            let media_path = if target.starts_with("media/") {
                format!("word/{}", target)
            } else {
                target.clone()
            };

            // Extract the filename
            let filename = media_path
                .rsplit('/')
                .next()
                .unwrap_or("unknown")
                .to_string();

            // Determine content type from extension
            let content_type = guess_content_type(&filename);

            // Read the media file
            let data = match read_archive_bytes(archive, &media_path) {
                Ok(d) => d,
                Err(_) => continue, // Skip if file not found
            };

            media.push(MediaFile {
                filename,
                data,
                content_type,
            });
        }
    }

    Ok(())
}

/// Read a file from the ZIP archive as bytes
fn read_archive_bytes<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<u8>> {
    let mut file = archive
        .by_name(path)
        .map_err(|e| Error::Zip(format!("Failed to read {}: {}", path, e)))?;
    let mut data = Vec::new();
    file.read_to_end(&mut data).map_err(Error::Io)?;
    Ok(data)
}

/// Guess content type from filename extension
pub fn guess_content_type(filename: &str) -> String {
    let ext = filename.rsplit('.').next().unwrap_or("").to_lowercase();

    match ext.as_str() {
        "png" => "image/png".to_string(),
        "jpg" | "jpeg" => "image/jpeg".to_string(),
        "gif" => "image/gif".to_string(),
        "bmp" => "image/bmp".to_string(),
        "svg" => "image/svg+xml".to_string(),
        "emf" => "image/x-emf".to_string(),
        _ => "application/octet-stream".to_string(),
    }
}

/// Extract tab stops from the Header and Footer styles in the template's styles.xml
///
/// Returns (header_tabs, footer_tabs) where each is a Vec of (position, alignment).
fn extract_style_tabs<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> Result<(Vec<(u32, String)>, Vec<(u32, String)>)> {
    let styles_xml = match archive.by_name("word/styles.xml") {
        Ok(mut file) => {
            let mut content = String::new();
            file.read_to_string(&mut content).map_err(Error::Io)?;
            content
        }
        Err(_) => return Ok((Vec::new(), Vec::new())),
    };

    let header_tabs = extract_tabs_for_style(&styles_xml, "Header");
    let footer_tabs = extract_tabs_for_style(&styles_xml, "Footer");

    Ok((header_tabs, footer_tabs))
}

/// Extract tab stops from a specific style definition in styles.xml
fn extract_tabs_for_style(styles_xml: &str, style_id: &str) -> Vec<(u32, String)> {
    // Find the style element with w:styleId="Header" or w:styleId="Footer"
    let pattern = format!(
        r#"<w:style[^>]*w:styleId="{}"[^>]*>.*?</w:style>"#,
        style_id
    );
    let style_regex = regex::Regex::new(&pattern).expect("style_regex should be valid");

    let style_match = match style_regex.find(styles_xml) {
        Some(m) => m.as_str(),
        None => return Vec::new(),
    };

    // Find <w:tabs>...</w:tabs> within the style
    let tabs_regex =
        regex::Regex::new(r"<w:tabs>(.*?)</w:tabs>").expect("tabs_regex should be valid");
    let tabs_content = match tabs_regex.captures(style_match) {
        Some(cap) => cap.get(1).unwrap().as_str(),
        None => return Vec::new(),
    };

    // Extract individual <w:tab> elements
    let tab_regex = regex::Regex::new(
        r#"<w:tab[^>]*w:val="([^"]*)"[^>]*w:pos="(\d+)"[^>]*/>"#,
    )
    .expect("tab_regex should be valid");

    let mut tabs = Vec::new();
    for cap in tab_regex.captures_iter(tabs_content) {
        let alignment = cap.get(1).unwrap().as_str().to_string();
        let position: u32 = cap.get(2).unwrap().as_str().parse().unwrap_or(0);
        tabs.push((position, alignment));
    }

    tabs
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_footer_template_default() {
        let template = HeaderFooterTemplate::default();
        assert!(template.is_empty());
        assert!(template.all_placeholders().is_empty());
    }

    #[test]
    fn test_extract_placeholders_from_xml() {
        let xml = r#"<w:p><w:r><w:t>{{title}} by {{author}}</w:t></w:r></w:p>"#;
        let placeholders = extract_placeholders_from_xml(xml);
        assert_eq!(placeholders, vec!["title", "author"]);
    }

    #[test]
    fn test_extract_placeholders_empty() {
        let xml = r#"<w:p><w:r><w:t>No placeholders here</w:t></w:r></w:p>"#;
        let placeholders = extract_placeholders_from_xml(xml);
        assert!(placeholders.is_empty());
    }

    #[test]
    fn test_guess_content_type() {
        assert_eq!(guess_content_type("image.png"), "image/png");
        assert_eq!(guess_content_type("photo.jpg"), "image/jpeg");
        assert_eq!(guess_content_type("animation.gif"), "image/gif");
        assert_eq!(guess_content_type("drawing.svg"), "image/svg+xml");
        assert_eq!(guess_content_type("metafile.emf"), "image/x-emf");
        assert_eq!(
            guess_content_type("unknown.xyz"),
            "application/octet-stream"
        );
    }

    #[test]
    fn test_find_target_by_rid() {
        let rels_xml = r#"
            <Relationships>
                <Relationship Id="rId1" Type="..." Target="header1.xml"/>
                <Relationship Id="rId2" Type="..." Target="header2.xml"/>
            </Relationships>
        "#;

        assert_eq!(
            find_target_by_rid(rels_xml, "rId1"),
            Some("header1.xml".to_string())
        );
        assert_eq!(
            find_target_by_rid(rels_xml, "rId2"),
            Some("header2.xml".to_string())
        );
        assert_eq!(find_target_by_rid(rels_xml, "rId99"), None);
    }
}
