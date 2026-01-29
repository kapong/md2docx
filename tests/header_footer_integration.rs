//! Integration tests for header/footer template extraction and rendering

use md2docx::template::extract::header_footer::{
    extract_placeholders_from_xml, guess_content_type, HeaderFooterContent, HeaderFooterTemplate,
};
use md2docx::template::render::header_footer::{render_header_footer, HeaderFooterContext};
use std::collections::HashMap;

#[test]
fn test_extract_placeholders_from_xml() {
    let xml = r#"<w:hdr><w:p><w:r><w:t>{{title}} by {{author}}</w:t></w:r></w:p></w:hdr>"#;
    let placeholders = extract_placeholders_from_xml(xml);
    assert_eq!(placeholders, vec!["title", "author"]);
}

#[test]
fn test_extract_placeholders_empty() {
    let xml = r#"<w:hdr><w:p><w:r><w:t>No placeholders here</w:t></w:r></w:p></w:hdr>"#;
    let placeholders = extract_placeholders_from_xml(xml);
    assert!(placeholders.is_empty());
}

#[test]
fn test_guess_content_type() {
    assert_eq!(guess_content_type("image.png"), "image/png");
    assert_eq!(guess_content_type("photo.jpg"), "image/jpeg");
    assert_eq!(guess_content_type("animation.gif"), "image/gif");
    assert_eq!(guess_content_type("drawing.svg"), "image/svg+xml");
    assert_eq!(
        guess_content_type("unknown.xyz"),
        "application/octet-stream"
    );
}

#[test]
fn test_header_footer_template_default() {
    let template = HeaderFooterTemplate::default();
    assert!(template.is_empty());
    assert!(template.all_placeholders().is_empty());
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

    let xml_str = String::from_utf8_lossy(&result.xml);
    assert!(xml_str.contains("Test Document"));
    assert!(xml_str.contains("PAGE"));
    assert!(!xml_str.contains("{{title}}"));
    assert!(!xml_str.contains("{{page}}"));
}

#[test]
fn test_render_header_footer_with_all_placeholders() {
    let content = HeaderFooterContent {
        raw_xml: r#"<w:hdr><w:p><w:r><w:t>{{title}} - {{subtitle}} by {{author}} on {{date}} - Page {{page}} of {{numpages}} - {{chapter}}</w:t></w:r></w:p></w:hdr>"#.to_string(),
        placeholders: vec![
            "title".to_string(),
            "subtitle".to_string(),
            "author".to_string(),
            "date".to_string(),
            "page".to_string(),
            "numpages".to_string(),
            "chapter".to_string(),
        ],
        rel_id_map: HashMap::new(),
    };

    let ctx = HeaderFooterContext::new("My Title", "John Doe")
        .with_subtitle("A Subtitle")
        .with_date("2025-01-29");

    let result = render_header_footer(&content, &ctx, 10, &[]).unwrap();

    let xml_str = String::from_utf8_lossy(&result.xml);
    assert!(xml_str.contains("My Title"));
    assert!(xml_str.contains("A Subtitle"));
    assert!(xml_str.contains("John Doe"));
    assert!(xml_str.contains("2025-01-29"));
    assert!(xml_str.contains("PAGE"));
    assert!(xml_str.contains("NUMPAGES"));
    assert!(xml_str.contains("STYLEREF"));
    assert!(!xml_str.contains("{{"));
}

#[test]
fn test_render_header_footer_xml_escaping() {
    let content = HeaderFooterContent {
        raw_xml: r#"<w:hdr><w:p><w:r><w:t>{{title}}</w:t></w:r></w:p></w:hdr>"#.to_string(),
        placeholders: vec!["title".to_string()],
        rel_id_map: HashMap::new(),
    };

    let ctx = HeaderFooterContext::new("Title & <Tag>", "Author");

    let result = render_header_footer(&content, &ctx, 10, &[]).unwrap();

    let xml_str = String::from_utf8_lossy(&result.xml);
    assert!(xml_str.contains("Title &amp; &lt;Tag&gt;"));
    assert!(!xml_str.contains("Title & <Tag>"));
}

#[test]
fn test_header_footer_context_builder() {
    let ctx = HeaderFooterContext::new("Title", "Author")
        .with_subtitle("Subtitle")
        .with_date("2025-01-29");

    assert_eq!(ctx.title, "Title");
    assert_eq!(ctx.author, "Author");
    assert_eq!(ctx.subtitle, "Subtitle");
    assert_eq!(ctx.date, "2025-01-29");
}
