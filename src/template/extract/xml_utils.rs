//! Shared XML parsing utilities for template extraction

/// Default values for run properties extraction
#[derive(Debug, Clone)]
pub struct RunPropertiesDefaults {
    pub font_family: &'static str,
    pub font_size: u32, // in half-points (22 = 11pt)
    pub font_color: &'static str,
}

impl Default for RunPropertiesDefaults {
    fn default() -> Self {
        Self {
            font_family: "Calibri",
            font_size: 22,
            font_color: "#000000",
        }
    }
}

impl RunPropertiesDefaults {
    /// Thai document defaults (TH Sarabun New, 12pt)
    pub fn thai() -> Self {
        Self {
            font_family: "TH Sarabun New",
            font_size: 24,
            font_color: "#000000",
        }
    }
}

/// Extracted run properties from OOXML
#[derive(Debug, Clone)]
pub struct RunProperties {
    pub font_family: String,
    pub font_size: u32,
    pub font_color: String,
    pub bold: bool,
    pub italic: bool,
}

/// Extract run properties (font, size, color, bold, italic) from OOXML run properties XML
///
/// # Arguments
/// * `rpr_content` - The content inside <w:rPr>...</w:rPr>
/// * `defaults` - Default values to use when properties are not specified
///
/// # Returns
/// RunProperties struct with extracted or default values
pub fn extract_run_properties(
    rpr_content: &str,
    defaults: &RunPropertiesDefaults,
) -> RunProperties {
    // Extract font family - check w:ascii first, then w:cs for complex script fonts
    let font_family = extract_attribute(rpr_content, "w:ascii=")
        .or_else(|| extract_attribute(rpr_content, "w:cs="))
        .unwrap_or_else(|| defaults.font_family.to_string());

    // Extract font size (w:sz or w:szCs for complex script)
    let font_size = extract_attribute(rpr_content, "w:val=")
        .and_then(|s| s.parse::<u32>().ok())
        .unwrap_or(defaults.font_size);

    // Extract font color
    let font_color = extract_attribute(rpr_content, "w:val=")
        .map(|c| {
            if c.len() == 6 && c.chars().all(|ch| ch.is_ascii_hexdigit()) {
                format!("#{}", c)
            } else {
                c
            }
        })
        .unwrap_or_else(|| defaults.font_color.to_string());

    // Check for bold (w:b or w:bCs)
    let bold = rpr_content.contains("<w:b/>")
        || rpr_content.contains("<w:b ")
        || rpr_content.contains("<w:bCs/>");

    // Check for italic (w:i or w:iCs)
    let italic = rpr_content.contains("<w:i/>")
        || rpr_content.contains("<w:i ")
        || rpr_content.contains("<w:iCs/>");

    RunProperties {
        font_family,
        font_size,
        font_color,
        bold,
        italic,
    }
}

/// Extract an attribute value from XML by attribute name
///
/// Finds the attribute in the XML string and extracts its quoted value.
///
/// # Arguments
/// * `xml` - The XML string to search
/// * `attr_name` - The attribute name including the equals sign (e.g., "w:val=")
///
/// # Returns
/// The attribute value if found, None otherwise
///
/// # Example
/// ```ignore
/// let xml = r#"<w:sz w:val="24"/>"#;
/// assert_eq!(extract_attribute(xml, "w:val="), Some("24".to_string()));
/// ```
pub fn extract_attribute(xml: &str, attr_name: &str) -> Option<String> {
    if let Some(pos) = xml.find(attr_name) {
        let start = pos + attr_name.len();
        let rest = &xml[start..];
        if let Some(quote_pos) = rest.find('"') {
            let after_quote = &rest[quote_pos + 1..];
            if let Some(end_quote) = after_quote.find('"') {
                return Some(after_quote[..end_quote].to_string());
            }
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_extract_attribute() {
        let xml = r#"<w:sz w:val="24"/>"#;
        assert_eq!(extract_attribute(xml, "w:val="), Some("24".to_string()));
    }

    #[test]
    fn test_extract_attribute_not_found() {
        let xml = r#"<w:sz w:val="24"/>"#;
        assert_eq!(extract_attribute(xml, "w:other="), None);
    }

    #[test]
    fn test_extract_attribute_complex() {
        let xml = r#"<w:rFonts w:ascii="Calibri" w:hAnsi="Arial"/>"#;
        assert_eq!(
            extract_attribute(xml, "w:ascii="),
            Some("Calibri".to_string())
        );
        assert_eq!(
            extract_attribute(xml, "w:hAnsi="),
            Some("Arial".to_string())
        );
    }
}
