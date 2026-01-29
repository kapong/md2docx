//! Template rendering for applying extracted templates to documents
//!
//! This module provides functionality to render extracted templates
//! with actual data and apply them to document generation.

pub mod cover;
pub mod header_footer;
pub mod table;

pub use cover::render_cover;
pub use header_footer::{
    render_default_footer, render_default_header, render_first_page_footer,
    render_first_page_header, render_header_footer, HeaderFooterContext, RenderedHeaderFooter,
};
pub use table::render_table_with_template;
