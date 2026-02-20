mod content_types;
mod doc_props;
mod document;
mod endnotes;
mod footer;
mod footnotes;
mod header;
mod latent_styles;
pub(crate) mod numbering;
mod rels;
mod styles;

// Re-export types for internal use within the crate
pub(crate) use content_types::ContentTypes;
pub(crate) use doc_props::*;
pub(crate) use document::{
    DocElement, DocumentXml, HeaderFooterRefs, Hyperlink, ImageBorderEffect, ImageEffectExtent,
    ImageElement, ImageShadowEffect, PageLayout, ParagraphChild, Table, TableCellElement, TableRow,
    TableWidth,
};
pub(crate) use endnotes::*;
pub(crate) use footer::*;
pub(crate) use header::*;
pub(crate) use rels::Relationships;
pub(crate) use styles::{
    generate_font_table_xml, generate_settings_xml, generate_theme_xml, generate_web_settings_xml,
    StylesDocument,
};

// Public API exports
pub use document::{Paragraph, Run, TabStop};
pub use footer::FooterConfig;
pub use footnotes::FootnotesXml;
pub use header::{HeaderConfig, HeaderFooterField};
pub use styles::{FontConfig, Language};
