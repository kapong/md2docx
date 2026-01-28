//! Demo showing generated OOXML XML output

use md2docx::docx::ooxml::{ContentTypes, Relationships};

fn main() {
    println!("=== [Content_Types].xml ===\n");
    let ct = ContentTypes::new();
    let xml = ct.to_xml().unwrap();
    println!("{}", String::from_utf8(xml).unwrap());

    println!("\n\n=== _rels/.rels ===\n");
    let root_rels = Relationships::root_rels();
    let xml = root_rels.to_xml().unwrap();
    println!("{}", String::from_utf8(xml).unwrap());

    println!("\n\n=== word/_rels/document.xml.rels ===\n");
    let doc_rels = Relationships::document_rels();
    let xml = doc_rels.to_xml().unwrap();
    println!("{}", String::from_utf8(xml).unwrap());

    println!("\n\n=== Extended example with images, headers, footers ===\n");

    let mut ct = ContentTypes::new();
    ct.add_image_extension("png", "image/png");
    ct.add_image_extension("jpeg", "image/jpeg");
    ct.add_numbering();
    ct.add_header(1);
    ct.add_footer(1);

    println!("=== [Content_Types].xml (extended) ===\n");
    let xml = ct.to_xml().unwrap();
    println!("{}", String::from_utf8(xml).unwrap());

    let mut rels = Relationships::document_rels();
    rels.add_numbering();
    rels.add_header(1);
    rels.add_footer(1);
    rels.add_image("logo.png");
    rels.add_hyperlink("https://example.com");

    println!("\n=== word/_rels/document.xml.rels (extended) ===\n");
    let xml = rels.to_xml().unwrap();
    println!("{}", String::from_utf8(xml).unwrap());
}
