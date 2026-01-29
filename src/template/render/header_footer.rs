//! Header/Footer template rendering
//!
//! Renders header/footer templates by:
//! - Replacing static placeholders ({{title}}, {{author}}, {{date}}, {{subtitle}}) with values
//! - Converting dynamic placeholders to Word fields:
//!   - {{page}} -> PAGE field
//!   - {{numpages}} -> NUMPAGES field
//!   - {{chapter}} -> STYLEREF "Heading 1" field

use crate::error::Result;
use crate::template::extract::header_footer::{HeaderFooterTemplate, MediaFile};
use regex::Regex;
use std::collections::HashMap;

/// Context for placeholder replacement
/// Values come from the document section of md2docx.toml
#[derive(Debug, Clone, Default)]
pub struct HeaderFooterContext {
    pub title: String,
    pub subtitle: String,
    pub author: String,
    pub date: String,
}

impl HeaderFooterContext {
    /// Create a new context with required fields
    pub fn new(title: impl Into<String>, author: impl Into<String>) -> Self {
        Self {
            title: title.into(),
            author: author.into(),
            ..Default::default()
        }
    }

    /// Set the subtitle
    pub fn with_subtitle(mut self, subtitle: impl Into<String>) -> Self {
        self.subtitle = subtitle.into();
        self
    }

    /// Set the date
    pub fn with_date(mut self, date: impl Into<String>) -> Self {
        self.date = date.into();
        self
    }
}

/// Rendered header/footer ready to be written to DOCX
#[derive(Debug, Clone)]
pub struct RenderedHeaderFooter {
    /// The XML content with placeholders replaced
    pub xml: Vec<u8>,
    /// Media files that need to be included (with remapped rIds)
    pub media: Vec<(String, MediaFile)>, // (new_rId, media_file)
}

/// Render a header/footer template
///
/// Static placeholders are replaced with values from context:
/// - {{title}} -> ctx.title
/// - {{subtitle}} -> ctx.subtitle
/// - {{author}} -> ctx.author
/// - {{date}} -> ctx.date
///
/// Dynamic placeholders are converted to Word fields:
/// - {{page}} -> PAGE field
/// - {{numpages}} -> NUMPAGES field
/// - {{chapter}} -> STYLEREF "Heading 1" field
pub fn render_header_footer(
    content: &crate::template::extract::header_footer::HeaderFooterContent,
    ctx: &HeaderFooterContext,
    rel_id_offset: u32,
    media_files: &[MediaFile],
) -> Result<RenderedHeaderFooter> {
    let mut xml = content.raw_xml.clone();

    // First, consolidate fragmented placeholders (Word often splits text across runs)
    xml = consolidate_fragmented_placeholders(&xml);

    // Replace static placeholders
    xml = xml.replace("{{title}}", &xml_escape(&ctx.title));
    xml = xml.replace("{{subtitle}}", &xml_escape(&ctx.subtitle));
    xml = xml.replace("{{author}}", &xml_escape(&ctx.author));
    xml = xml.replace("{{date}}", &xml_escape(&ctx.date));

    // Replace dynamic placeholders with Word fields
    xml = replace_page_placeholder(&xml);
    xml = replace_numpages_placeholder(&xml);
    xml = replace_chapter_placeholder(&xml);

    // Remap relationship IDs for media files and update XML
    let (media, rid_replacements) =
        remap_media_ids(&content.rel_id_map, rel_id_offset, media_files);

    // Replace old rIds with new rIds in the XML
    for (old_rid, new_rid) in &rid_replacements {
        // Replace r:embed="rIdX" patterns
        xml = xml.replace(
            &format!(r#"r:embed="{}""#, old_rid),
            &format!(r#"r:embed="{}""#, new_rid),
        );
        // Also replace r:id="rIdX" patterns (for hyperlinks in headers)
        xml = xml.replace(
            &format!(r#"r:id="{}""#, old_rid),
            &format!(r#"r:id="{}""#, new_rid),
        );
    }

    Ok(RenderedHeaderFooter {
        xml: xml.into_bytes(),
        media,
    })
}

/// Render the default header from a template
pub fn render_default_header(
    template: &HeaderFooterTemplate,
    ctx: &HeaderFooterContext,
    rel_id_offset: u32,
) -> Result<Option<RenderedHeaderFooter>> {
    match &template.default_header {
        Some(content) => {
            render_header_footer(content, ctx, rel_id_offset, &template.media).map(Some)
        }
        None => Ok(None),
    }
}

/// Render the default footer from a template
pub fn render_default_footer(
    template: &HeaderFooterTemplate,
    ctx: &HeaderFooterContext,
    rel_id_offset: u32,
) -> Result<Option<RenderedHeaderFooter>> {
    match &template.default_footer {
        Some(content) => {
            render_header_footer(content, ctx, rel_id_offset, &template.media).map(Some)
        }
        None => Ok(None),
    }
}

/// Render the first page header from a template
pub fn render_first_page_header(
    template: &HeaderFooterTemplate,
    ctx: &HeaderFooterContext,
    rel_id_offset: u32,
) -> Result<Option<RenderedHeaderFooter>> {
    match &template.first_page_header {
        Some(content) => {
            render_header_footer(content, ctx, rel_id_offset, &template.media).map(Some)
        }
        None => Ok(None),
    }
}

/// Render the first page footer from a template
pub fn render_first_page_footer(
    template: &HeaderFooterTemplate,
    ctx: &HeaderFooterContext,
    rel_id_offset: u32,
) -> Result<Option<RenderedHeaderFooter>> {
    match &template.first_page_footer {
        Some(content) => {
            render_header_footer(content, ctx, rel_id_offset, &template.media).map(Some)
        }
        None => Ok(None),
    }
}

/// Generate relationships XML for a header/footer
///
/// Creates a .rels file that maps relationship IDs to media file targets.
/// This is required for images in headers/footers to display correctly.
pub fn generate_header_footer_rels_xml(media: &[(String, MediaFile)]) -> Vec<u8> {
    generate_header_footer_rels_xml_with_prefix(media, "")
}

/// Generate relationships XML for a header/footer with optional filename prefix
///
/// The `prefix` is added to each media filename to avoid conflicts with images
/// from other templates (e.g., cover.docx vs header-footer.docx).
pub fn generate_header_footer_rels_xml_with_prefix(
    media: &[(String, MediaFile)],
    prefix: &str,
) -> Vec<u8> {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#,
    );

    for (r_id, media_file) in media {
        let filename = if prefix.is_empty() {
            media_file.filename.clone()
        } else {
            format!("{}{}", prefix, media_file.filename)
        };
        xml.push_str(&format!(
            r#"  <Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{}"/>
"#,
            r_id, filename
        ));
    }

    xml.push_str("</Relationships>");
    xml.into_bytes()
}

/// Escape XML special characters
fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Consolidate fragmented placeholders in Word XML
///
/// Word often splits text across multiple <w:r><w:t>...</w:t></w:r> elements,
/// especially when text is edited. For example, `{{numpages}}` might become:
/// ```xml
/// <w:r><w:t>prefix {{</w:t></w:r><w:r><w:t>numpages</w:t></w:r><w:r><w:t>}}</w:t></w:r>
/// ```
///
/// This function finds and consolidates such patterns back into continuous text.
fn consolidate_fragmented_placeholders(xml: &str) -> String {
    // List of placeholders we want to consolidate
    let placeholders = [
        "page", "numpages", "chapter", "title", "subtitle", "author", "date",
    ];

    let mut result = xml.to_string();

    for placeholder in placeholders {
        let full_placeholder = format!("{{{{{}}}}}", placeholder); // e.g., "{{page}}"

        // Already consolidated? Skip.
        if result.contains(&full_placeholder) {
            continue;
        }

        // Try multiple fragmentation patterns
        result = try_consolidate_three_run_pattern(&result, placeholder, &full_placeholder);
    }

    result
}

/// Try to consolidate the common 3-run fragmentation pattern:
/// Run 1: ends with "{{"
/// Run 2: contains placeholder name
/// Run 3: starts with "}}"
fn try_consolidate_three_run_pattern(xml: &str, placeholder: &str, full: &str) -> String {
    // Build a flexible regex that matches:
    // <w:r...><w:rPr>...</w:rPr><w:t...>...{{</w:t></w:r>
    // <w:r...><w:rPr>...</w:rPr><w:t>PLACEHOLDER</w:t></w:r>
    // <w:r...><w:rPr>...</w:rPr><w:t>}}...</w:t></w:r>
    //
    // Key: Match across all the XML tags between the fragments

    let pattern = format!(
        r#"(?s)(<w:r[^>]*>(?:\s*<w:rPr>.*?</w:rPr>)?(?:\s*<w:tab/>)?\s*<w:t[^>]*>)([^<]*)\{{\{{\s*</w:t>\s*</w:r>(?:\s*<w:proofErr[^/]*/>)?\s*<w:r[^>]*>\s*(?:<w:rPr>.*?</w:rPr>)?\s*<w:t[^>]*>\s*{}\s*</w:t>\s*</w:r>(?:\s*<w:proofErr[^/]*/>)?\s*<w:r[^>]*>\s*(?:<w:rPr>.*?</w:rPr>)?\s*<w:t[^>]*>\s*\}}\}}([^<]*)</w:t>\s*</w:r>"#,
        regex::escape(placeholder)
    );

    if let Ok(re) = Regex::new(&pattern) {
        re.replace_all(xml, |caps: &regex::Captures| {
            // caps[1] = opening of first run up to <w:t...>
            // caps[2] = prefix text before {{
            // caps[3] = suffix text after }}
            let opening = caps.get(1).map(|m| m.as_str()).unwrap_or("<w:r><w:t>");
            let prefix = caps.get(2).map(|m| m.as_str()).unwrap_or("");
            let suffix = caps.get(3).map(|m| m.as_str()).unwrap_or("");

            // Reconstruct as a single run
            format!("{}{}{}{}</w:t></w:r>", opening, prefix, full, suffix)
        })
        .to_string()
    } else {
        xml.to_string()
    }
}

/// Replace {{page}} placeholder with Word PAGE field
fn replace_page_placeholder(xml: &str) -> String {
    // Note: consolidate_fragmented_placeholders() is called before this,
    // so placeholders should already be in a single run by this point.
    let page_field = r#"</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">"#;
    xml.replace("{{page}}", page_field)
}

/// Replace {{numpages}} placeholder with Word NUMPAGES field
fn replace_numpages_placeholder(xml: &str) -> String {
    let numpages_field = r#"</w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">"#;
    xml.replace("{{numpages}}", numpages_field)
}

/// Replace {{chapter}} placeholder with Word STYLEREF field
fn replace_chapter_placeholder(xml: &str) -> String {
    let chapter_field = r#"</w:t></w:r><w:fldSimple w:instr="STYLEREF &quot;Heading 1&quot; \* MERGEFORMAT"><w:r><w:rPr><w:noProof/></w:rPr><w:t>Chapter</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">"#;
    xml.replace("{{chapter}}", chapter_field)
}

/// Remap relationship IDs for media files
///
/// This ensures that when we embed media files in the final document,
/// they get new rIds that don't conflict with existing relationships.
///
/// Returns:
/// - Vec of (new_rId, MediaFile) tuples for the rels file
/// - Vec of (old_rId, new_rId) tuples for XML replacement
fn remap_media_ids(
    rel_id_map: &HashMap<String, String>,
    offset: u32,
    media_files: &[MediaFile],
) -> (Vec<(String, MediaFile)>, Vec<(String, String)>) {
    let mut media = Vec::new();
    let mut rid_replacements = Vec::new();

    // Sort rIds to ensure consistent ordering
    let mut r_ids: Vec<_> = rel_id_map.keys().collect();
    r_ids.sort();

    // Create a map from target path to MediaFile
    let mut media_by_target: HashMap<String, &MediaFile> = HashMap::new();
    for mf in media_files {
        media_by_target.insert(mf.filename.clone(), mf);
    }

    for (i, old_r_id) in r_ids.iter().enumerate() {
        let new_r_id = format!("rId{}", offset + i as u32);

        // Look up the target path for this rId
        if let Some(target) = rel_id_map.get(*old_r_id) {
            // Extract filename from target path (e.g., "media/image1.png" -> "image1.png")
            let filename = target.rsplit('/').next().unwrap_or(target).to_string();

            // Find the corresponding media file
            if let Some(mf) = media_by_target.get(&filename) {
                media.push((new_r_id.clone(), (*mf).clone()));
                rid_replacements.push(((*old_r_id).clone(), new_r_id));
            }
        }
    }

    (media, rid_replacements)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::template::extract::header_footer::HeaderFooterContent;

    #[test]
    fn test_xml_escape() {
        assert_eq!(xml_escape("Hello & World"), "Hello &amp; World");
        assert_eq!(xml_escape("<tag>"), "&lt;tag&gt;");
        assert_eq!(xml_escape("\"quoted\""), "&quot;quoted&quot;");
        assert_eq!(xml_escape("'apostrophe'"), "&apos;apostrophe&apos;");
    }

    #[test]
    fn test_replace_page_placeholder() {
        let xml = r#"<w:p><w:r><w:t>Page {{page}}</w:t></w:r></w:p>"#;
        let result = replace_page_placeholder(xml);
        assert!(result.contains("PAGE"));
        assert!(!result.contains("{{page}}"));
    }

    #[test]
    fn test_replace_numpages_placeholder() {
        let xml = r#"<w:p><w:r><w:t>Total: {{numpages}}</w:t></w:r></w:p>"#;
        let result = replace_numpages_placeholder(xml);
        assert!(result.contains("NUMPAGES"));
        assert!(!result.contains("{{numpages}}"));
    }

    #[test]
    fn test_replace_chapter_placeholder() {
        let xml = r#"<w:p><w:r><w:t>Chapter: {{chapter}}</w:t></w:r></w:p>"#;
        let result = replace_chapter_placeholder(xml);
        assert!(result.contains("STYLEREF"));
        assert!(!result.contains("{{chapter}}"));
    }

    #[test]
    fn test_header_footer_context() {
        let ctx = HeaderFooterContext::new("My Title", "John Doe")
            .with_subtitle("A Subtitle")
            .with_date("2025-01-29");

        assert_eq!(ctx.title, "My Title");
        assert_eq!(ctx.author, "John Doe");
        assert_eq!(ctx.subtitle, "A Subtitle");
        assert_eq!(ctx.date, "2025-01-29");
    }

    #[test]
    fn test_render_header_footer() {
        let content = HeaderFooterContent {
            raw_xml: r#"<w:hdr><w:p><w:r><w:t>{{title}} - {{page}}</w:t></w:r></w:p></w:hdr>"#
                .to_string(),
            placeholders: vec!["title".to_string(), "page".to_string()],
            rel_id_map: HashMap::new(),
        };

        let ctx = HeaderFooterContext::new("Test Document", "Author");

        let result = render_header_footer(&content, &ctx, 10, &[]).unwrap();

        assert!(String::from_utf8_lossy(&result.xml).contains("Test Document"));
        assert!(String::from_utf8_lossy(&result.xml).contains("PAGE"));
        assert!(!String::from_utf8_lossy(&result.xml).contains("{{title}}"));
        assert!(!String::from_utf8_lossy(&result.xml).contains("{{page}}"));
    }

    #[test]
    fn test_render_default_header_none() {
        let template = HeaderFooterTemplate::default();
        let ctx = HeaderFooterContext::default();

        let result = render_default_header(&template, &ctx, 0).unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_render_default_footer_none() {
        let template = HeaderFooterTemplate::default();
        let ctx = HeaderFooterContext::default();

        let result = render_default_footer(&template, &ctx, 0).unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_consolidate_fragmented_numpages() {
        // Simulate Word's fragmented XML for {{numpages}}
        let fragmented = r#"<w:r><w:t xml:space="preserve"> / {{</w:t></w:r><w:r><w:t>numpages</w:t></w:r><w:r><w:t>}}</w:t></w:r>"#;

        let result = consolidate_fragmented_placeholders(fragmented);

        // Should be consolidated into a single run
        assert!(
            result.contains("{{numpages}}"),
            "Should contain consolidated placeholder. Got: {}",
            result
        );
    }

    #[test]
    fn test_consolidate_fragmented_numpages_with_attrs() {
        // Simulate the exact Word XML structure with rsidR attributes and rPr elements
        let fragmented = r#"<w:r><w:t xml:space="preserve"> / {{</w:t></w:r><w:r w:rsidR="00E63927"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>numpages</w:t></w:r><w:r w:rsidR="00F60378"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>}}</w:t></w:r>"#;

        let result = consolidate_fragmented_placeholders(fragmented);

        // Should be consolidated into a single run
        assert!(
            result.contains("{{numpages}}"),
            "Should contain consolidated placeholder. Got: {}",
            result
        );
        // The prefix " / " should be preserved
        assert!(
            result.contains(" / {{numpages}}"),
            "Should preserve prefix text. Got: {}",
            result
        );
    }

    #[test]
    fn test_consolidate_fragmented_page() {
        // Simulate fragmented {{page}}
        let fragmented =
            r#"<w:r><w:t>{{</w:t></w:r><w:r><w:t>page</w:t></w:r><w:r><w:t>}}</w:t></w:r>"#;

        let result = consolidate_fragmented_placeholders(fragmented);

        assert!(
            result.contains("{{page}}"),
            "Should contain consolidated placeholder. Got: {}",
            result
        );
    }

    #[test]
    fn test_already_consolidated_placeholder() {
        // Already consolidated - should remain unchanged (minus any normalization)
        let consolidated = r#"<w:r><w:t>Page: {{page}} of {{numpages}}</w:t></w:r>"#;

        let result = consolidate_fragmented_placeholders(consolidated);

        assert!(result.contains("{{page}}"));
        assert!(result.contains("{{numpages}}"));
    }

    #[test]
    fn test_consolidate_real_template_pattern() {
        // This is the exact pattern from the real template footer2.xml
        // where {{page}} is intact but {{numpages}} is fragmented across 3 runs
        let fragmented = r#"<w:r w:rsidR="00F60378"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>{{page}} / {{</w:t></w:r><w:r w:rsidR="00E63927"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>numpages</w:t></w:r><w:r w:rsidR="00F60378"><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>}}</w:t></w:r>"#;

        let result = consolidate_fragmented_placeholders(fragmented);

        // Should contain both placeholders
        assert!(
            result.contains("{{page}}"),
            "Should contain {{page}}. Got: {}",
            result
        );
        assert!(
            result.contains("{{numpages}}"),
            "Should contain consolidated {{numpages}}. Got: {}",
            result
        );
    }

    #[test]
    fn test_generate_header_footer_rels_xml() {
        use crate::template::extract::header_footer::MediaFile;

        let media = vec![
            (
                "rId1".to_string(),
                MediaFile {
                    filename: "image1.png".to_string(),
                    data: vec![1, 2, 3],
                    content_type: "image/png".to_string(),
                },
            ),
            (
                "rId2".to_string(),
                MediaFile {
                    filename: "logo.jpg".to_string(),
                    data: vec![4, 5, 6],
                    content_type: "image/jpeg".to_string(),
                },
            ),
        ];

        let xml = generate_header_footer_rels_xml(&media);
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<?xml version=\"1.0\""));
        assert!(xml_str.contains("<Relationships"));
        assert!(xml_str.contains("rId1"));
        assert!(xml_str.contains("rId2"));
        assert!(xml_str.contains("media/image1.png"));
        assert!(xml_str.contains("media/logo.jpg"));
        assert!(xml_str
            .contains("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"));
    }

    #[test]
    fn test_generate_header_footer_rels_xml_empty() {
        let media: Vec<(String, MediaFile)> = vec![];
        let xml = generate_header_footer_rels_xml(&media);
        let xml_str = String::from_utf8(xml).unwrap();

        assert!(xml_str.contains("<?xml version=\"1.0\""));
        assert!(xml_str.contains("<Relationships"));
        assert!(xml_str.contains("</Relationships>"));
        // Should not contain any Relationship elements
        assert!(!xml_str.contains("<Relationship Id="));
    }
}
