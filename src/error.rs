//! Error types for md2docx

use thiserror::Error;

/// Main error type for md2docx operations
#[derive(Debug, Error)]
pub enum Error {
    /// Failed to parse markdown content
    #[error("Failed to parse markdown: {0}")]
    Parse(String),

    /// IO error (file read/write)
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    /// XML generation/parsing error
    #[error("XML error: {0}")]
    Xml(String),

    /// ZIP archive error
    #[error("ZIP error: {0}")]
    Zip(String),

    /// Configuration error
    #[error("Configuration error: {0}")]
    Config(String),

    /// Template error (missing styles, invalid template)
    #[error("Template error: {0}")]
    Template(String),

    /// Image processing error
    #[error("Image error: {0}")]
    Image(String),

    /// Mermaid rendering error
    #[error("Mermaid error: {0}")]
    Mermaid(String),

    /// Git diff error
    #[error("Git error: {0}")]
    Git(String),

    /// Include directive error
    #[error("Include error: {0}")]
    Include(String),

    /// Regex compilation error
    #[error("Regex error: {0}")]
    Regex(String),

    /// UTF-8 conversion error
    #[error("UTF-8 error: {0}")]
    Utf8(String),

    /// Template parsing error
    #[error("Template parse error: {0}")]
    TemplateParse(String),

    /// Feature not implemented yet
    #[error("Not implemented: {0}")]
    NotImplemented(String),
}

/// Result type alias for md2docx operations
pub type Result<T> = std::result::Result<T, Error>;

// Implement From for zip::result::ZipError
impl From<zip::result::ZipError> for Error {
    fn from(err: zip::result::ZipError) -> Self {
        Error::Zip(err.to_string())
    }
}

// Implement From for quick_xml::Error
impl From<quick_xml::Error> for Error {
    fn from(err: quick_xml::Error) -> Self {
        Error::Xml(err.to_string())
    }
}

// Implement From for regex::Error
impl From<regex::Error> for Error {
    fn from(err: regex::Error) -> Self {
        Error::Regex(err.to_string())
    }
}

// Implement From for std::string::FromUtf8Error
impl From<std::string::FromUtf8Error> for Error {
    fn from(err: std::string::FromUtf8Error) -> Self {
        Error::Utf8(err.to_string())
    }
}
