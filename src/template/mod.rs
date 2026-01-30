//! Directory-based template system for md2docx
//!
//! This module provides a template system where users can design templates
//! visually in Microsoft Word by creating separate DOCX files for different
//! components:
//!
//! - `cover.docx` - Cover page design with placeholders like {{title}}, {{author}}
//! - `table.docx` - Table style example with header, odd/even rows, first column
//! - `image.docx` - Image caption style
//! - `header-footer.docx` - Header/footer with placeholders
//!
//! # Example Template Directory Structure
//!
//! ```text
//! my-template/
//! ├── cover.docx          # Cover page design
//! ├── table.docx          # Table style example
//! ├── image.docx          # Image caption style
//! └── header-footer.docx  # Header/footer placeholders
//! ```
//!
//! # Usage
//!
//! ```rust,no_run
//! use md2docx::template::TemplateDir;
//! use std::path::Path;
//!
//! let template = TemplateDir::load(Path::new("my-template")).unwrap();
//!
//! // Extract cover template
//! if let Some(cover) = template.extract_cover().unwrap() {
//!     // Use cover template...
//! }
//! ```

pub mod extract;
pub mod placeholder;
pub mod render;

pub use extract::{
    CoverElement, CoverTemplate, HeaderFooterContent, HeaderFooterTemplate, ImageTemplate,
    MediaFile, PageMargins, ShapeType, TableTemplate,
};
pub use placeholder::{
    extract_placeholders, has_placeholders, replace_placeholders, PlaceholderContext,
};

use crate::error::{Error, Result};
use std::path::{Path, PathBuf};

/// Represents a directory containing template DOCX files
#[derive(Debug, Clone)]
pub struct TemplateDir {
    /// Path to the template directory
    pub path: PathBuf,
}

impl TemplateDir {
    /// Load a template directory
    ///
    /// # Arguments
    /// * `path` - Path to the template directory
    ///
    /// # Returns
    /// A `TemplateDir` instance
    ///
    /// # Example
    /// ```rust,no_run
    /// use md2docx::template::TemplateDir;
    /// use std::path::Path;
    ///
    /// let template = TemplateDir::load(Path::new("my-template")).unwrap();
    /// ```
    pub fn load(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Err(Error::Template(format!(
                "Template directory does not exist: {}",
                path.display()
            )));
        }

        if !path.is_dir() {
            return Err(Error::Template(format!(
                "Template path is not a directory: {}",
                path.display()
            )));
        }

        Ok(Self {
            path: path.to_path_buf(),
        })
    }

    /// Check if a template file exists
    fn has_file(&self, filename: &str) -> bool {
        self.path.join(filename).exists()
    }

    /// Get the path to a template file
    fn file_path(&self, filename: &str) -> PathBuf {
        self.path.join(filename)
    }

    /// Extract cover template from `cover.docx`
    ///
    /// Returns `None` if cover.docx doesn't exist
    pub fn extract_cover(&self) -> Result<Option<CoverTemplate>> {
        if !self.has_file("cover.docx") {
            return Ok(None);
        }

        let path = self.file_path("cover.docx");
        extract::extract_cover(&path).map(Some)
    }

    /// Extract table template from `table.docx`
    ///
    /// Returns `None` if table.docx doesn't exist
    pub fn extract_table(&self) -> Result<Option<TableTemplate>> {
        if !self.has_file("table.docx") {
            return Ok(None);
        }

        let path = self.file_path("table.docx");
        extract::extract_table(&path).map(Some)
    }

    /// Extract image template from `image.docx`
    ///
    /// Returns `None` if image.docx doesn't exist
    pub fn extract_image(&self) -> Result<Option<ImageTemplate>> {
        if !self.has_file("image.docx") {
            return Ok(None);
        }

        let path = self.file_path("image.docx");
        extract::extract_image(&path).map(Some)
    }

    /// Extract header/footer template from `header-footer.docx`
    ///
    /// Returns `None` if header-footer.docx doesn't exist
    pub fn extract_header_footer(&self) -> Result<Option<HeaderFooterTemplate>> {
        if !self.has_file("header-footer.docx") {
            return Ok(None);
        }

        let path = self.file_path("header-footer.docx");
        extract::extract_header_footer(&path).map(Some)
    }

    /// Load all available templates
    ///
    /// Returns a `TemplateSet` containing all extracted templates
    pub fn load_all(&self) -> Result<TemplateSet> {
        Ok(TemplateSet {
            cover: self.extract_cover()?,
            table: self.extract_table()?,
            image: self.extract_image()?,
            header_footer: self.extract_header_footer()?,
        })
    }
}

/// Collection of all loaded templates
#[derive(Debug, Clone, Default)]
pub struct TemplateSet {
    pub(crate) cover: Option<CoverTemplate>,
    pub(crate) table: Option<TableTemplate>,
    pub(crate) image: Option<ImageTemplate>,
    pub(crate) header_footer: Option<HeaderFooterTemplate>,
}

impl TemplateSet {
    /// Check if any templates are loaded
    pub fn is_empty(&self) -> bool {
        self.cover.is_none()
            && self.table.is_none()
            && self.image.is_none()
            && self.header_footer.is_none()
    }

    /// Check if cover template is available
    pub fn has_cover(&self) -> bool {
        self.cover.is_some()
    }

    /// Check if table template is available
    pub fn has_table(&self) -> bool {
        self.table.is_some()
    }

    /// Check if image template is available
    pub fn has_image(&self) -> bool {
        self.image.is_some()
    }

    /// Check if header/footer template is available
    pub fn has_header_footer(&self) -> bool {
        self.header_footer.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use tempfile::TempDir;

    #[test]
    fn test_template_dir_load() {
        let temp_dir = TempDir::new().unwrap();
        let template_path = temp_dir.path();

        // Create empty template directory
        fs::create_dir_all(template_path).unwrap();

        let template = TemplateDir::load(template_path);
        assert!(template.is_ok());
    }

    #[test]
    fn test_template_dir_not_found() {
        let result = TemplateDir::load(Path::new("/nonexistent/template"));
        assert!(result.is_err());
    }

    #[test]
    fn test_template_dir_not_a_directory() {
        let temp_dir = TempDir::new().unwrap();
        let file_path = temp_dir.path().join("not-a-dir");
        fs::write(&file_path, "content").unwrap();

        let result = TemplateDir::load(&file_path);
        assert!(result.is_err());
    }

    #[test]
    fn test_has_file() {
        let temp_dir = TempDir::new().unwrap();
        let template_path = temp_dir.path();

        // Create a file
        fs::write(template_path.join("cover.docx"), "fake content").unwrap();

        let template = TemplateDir::load(template_path).unwrap();
        assert!(template.has_file("cover.docx"));
        assert!(!template.has_file("table.docx"));
    }

    #[test]
    fn test_template_set_empty() {
        let set = TemplateSet::default();
        assert!(set.is_empty());
        assert!(!set.has_cover());
        assert!(!set.has_table());
        assert!(!set.has_image());
        assert!(!set.has_header_footer());
    }
}
