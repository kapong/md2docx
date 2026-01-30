//! Template extraction from DOCX files
//!
//! This module provides functions to extract template styles and content
//! from DOCX files created in Microsoft Word.

pub mod cover;
pub mod header_footer;
pub mod image;
pub mod table;
pub mod xml_utils;

pub(crate) use xml_utils::{extract_attribute, extract_run_properties, RunPropertiesDefaults};

pub use cover::{CoverElement, CoverTemplate, PageMargins, ShapeType};
pub use header_footer::{HeaderFooterContent, HeaderFooterTemplate, MediaFile};
pub use image::{
    CaptionRun, EffectExtent, ImageBorder, ImageCaptionStyle, ImageShadow, ImageTemplate,
};
pub use table::{
    BorderStyle, BorderStyles, CellMargins, CellSpacing, CellStyle, RowStyle, TableCaptionStyle,
    TableTemplate,
};

use crate::error::Result;
use std::path::Path;

/// Extract cover template from a DOCX file
///
/// # Arguments
/// * `path` - Path to the cover.docx file
///
/// # Returns
/// The extracted `CoverTemplate`
pub fn extract_cover(path: &Path) -> Result<CoverTemplate> {
    cover::extract(path)
}

/// Extract table template from a DOCX file
///
/// # Arguments
/// * `path` - Path to the table.docx file
///
/// # Returns
/// The extracted `TableTemplate`
pub fn extract_table(path: &Path) -> Result<TableTemplate> {
    table::extract(path)
}

/// Extract image template from a DOCX file
///
/// # Arguments
/// * `path` - Path to the image.docx file
///
/// # Returns
/// The extracted `ImageTemplate`
pub fn extract_image(path: &Path) -> Result<ImageTemplate> {
    image::extract(path)
}

/// Extract header/footer template from a DOCX file
///
/// # Arguments
/// * `path` - Path to the header-footer.docx file
///
/// # Returns
/// The extracted `HeaderFooterTemplate`
pub fn extract_header_footer(path: &Path) -> Result<HeaderFooterTemplate> {
    header_footer::extract(path)
}
