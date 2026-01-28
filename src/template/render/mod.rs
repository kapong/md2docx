//! Template rendering for applying extracted templates to documents
//!
//! This module provides functionality to render extracted templates
//! with actual data and apply them to document generation.

pub mod cover;
pub mod table;

pub use cover::render_cover;
pub use table::render_table_with_template;
