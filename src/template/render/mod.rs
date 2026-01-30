//! Template rendering for applying extracted templates to documents
//!
//! This module provides functionality to render extracted templates
//! with actual data and apply them to document generation.

pub mod cover;
pub mod header_footer;
pub mod table;

pub use table::render_table_with_template;
