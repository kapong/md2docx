//! Image template extraction from DOCX files
//!
//! Extracts image styling and caption styling from a DOCX file.
//! The file should contain a sample image with a caption using placeholders.

use super::{extract_attribute, extract_run_properties, RunPropertiesDefaults};
use crate::error::{Error, Result};
use std::io::Read;
use std::path::Path;
use zip::ZipArchive;

/// Image border properties
#[derive(Debug, Clone, Default)]
pub struct ImageBorder {
    /// Fill type: "solid", "gradient", "none"
    pub fill_type: String,
    /// Color value (hex like "#156082" or scheme name like "accent1")
    pub color: String,
    /// Whether color is a scheme color (theme-based)
    pub is_scheme_color: bool,
    /// Border width in EMUs (None = default)
    pub width: Option<u32>,
}

/// Image shadow effect
#[derive(Debug, Clone)]
pub struct ImageShadow {
    /// Blur radius in EMUs (e.g., 190500 = 15pt)
    pub blur_radius: u32,
    /// Shadow distance in EMUs (e.g., 228600 = 18pt)
    pub distance: u32,
    /// Direction in 60000ths of a degree (e.g., 2700000 = 270°)
    pub direction: u32,
    /// Alignment: "ctr", "tl", "tr", "bl", "br", etc.
    pub alignment: String,
    /// Shadow color (hex, e.g., "000000")
    pub color: String,
    /// Opacity in thousandths (30000 = 30%)
    pub alpha: u32,
}

impl Default for ImageShadow {
    fn default() -> Self {
        Self {
            blur_radius: 0,
            distance: 0,
            direction: 0,
            alignment: "ctr".to_string(),
            color: "000000".to_string(),
            alpha: 100000, // 100%
        }
    }
}

/// Effect extent (space for shadows/effects)
#[derive(Debug, Clone, Default)]
pub struct EffectExtent {
    pub left: u32, // EMUs
    pub top: u32,
    pub right: u32,
    pub bottom: u32,
}

/// Caption run formatting (prefix vs text may differ)
#[derive(Debug, Clone)]
pub struct CaptionRun {
    /// Placeholder type: "prefix", "number", "text"
    pub placeholder: String,
    /// Font family
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Whether bold
    pub bold: bool,
    /// Whether italic
    pub italic: bool,
    /// Font color (hex)
    pub font_color: String,
}

impl Default for CaptionRun {
    fn default() -> Self {
        Self {
            placeholder: "text".to_string(),
            font_family: "Calibri".to_string(),
            font_size: 22,
            bold: false,
            italic: false,
            font_color: "000000".to_string(),
        }
    }
}

/// Represents an extracted image template
#[derive(Debug, Clone)]
pub struct ImageTemplate {
    /// Caption style
    pub caption: ImageCaptionStyle,
    /// Image border style
    pub border: Option<ImageBorder>,
    /// Shadow effect
    pub shadow: Option<ImageShadow>,
    /// Effect extent (for shadow space)
    pub effect_extent: EffectExtent,
    /// Default image width percentage (0-100)
    pub default_width_percent: u32,
    /// Image alignment: "left", "center", "right"
    pub alignment: String,
    /// Caption runs with individual formatting
    pub caption_runs: Vec<CaptionRun>,
    /// Whether to lock aspect ratio
    pub lock_aspect_ratio: bool,
}

/// Image caption style
#[derive(Debug, Clone)]
pub struct ImageCaptionStyle {
    /// Caption position: "top" or "bottom" (images usually have bottom)
    pub position: String,
    /// Caption prefix (e.g., "Figure" or "รูปที่")
    pub prefix: String,
    /// Font family
    pub font_family: String,
    /// Font size in half-points
    pub font_size: u32,
    /// Font color (hex)
    pub font_color: String,
    /// Whether text is bold
    pub bold: bool,
    /// Whether text is italic
    pub italic: bool,
    /// Text alignment
    pub alignment: String,
    /// Spacing before caption in twips
    pub spacing_before: u32,
    /// Spacing after caption in twips
    pub spacing_after: u32,
}

impl Default for ImageCaptionStyle {
    fn default() -> Self {
        Self {
            position: "bottom".to_string(),
            prefix: "Figure".to_string(),
            font_family: "Calibri".to_string(),
            font_size: 22, // 11pt
            font_color: "#4a5568".to_string(),
            bold: false,
            italic: true,
            alignment: "center".to_string(),
            spacing_before: 120, // 6pt
            spacing_after: 120,  // 6pt
        }
    }
}

impl Default for ImageTemplate {
    fn default() -> Self {
        Self {
            caption: ImageCaptionStyle::default(),
            border: None,
            shadow: None,
            effect_extent: EffectExtent::default(),
            default_width_percent: 80,
            alignment: "center".to_string(),
            caption_runs: Vec::new(),
            lock_aspect_ratio: true,
        }
    }
}

impl ImageTemplate {
    /// Create a new image template with default styles
    pub fn new() -> Self {
        Self::default()
    }

    /// Set caption prefix (e.g., "Figure" or "รูปที่")
    pub fn with_caption_prefix(mut self, prefix: impl Into<String>) -> Self {
        self.caption.prefix = prefix.into();
        self
    }

    /// Set caption position ("top" or "bottom")
    pub fn with_caption_position(mut self, position: impl Into<String>) -> Self {
        self.caption.position = position.into();
        self
    }

    /// Set default image width percentage
    pub fn with_default_width(mut self, percent: u32) -> Self {
        self.default_width_percent = percent.min(100);
        self
    }

    /// Format a caption with the given number and text
    ///
    /// # Arguments
    /// * `number` - The figure number (e.g., "1.2")
    /// * `text` - The caption text
    ///
    /// # Returns
    /// Formatted caption string (e.g., "Figure 1.2: Caption text")
    pub fn format_caption(&self, number: &str, text: &str) -> String {
        format!("{} {}: {}", self.caption.prefix, number, text)
    }
}

/// Extract image template from a DOCX file
///
/// This function reads a DOCX file and extracts image styling and caption styling.
///
/// # Arguments
/// * `path` - Path to the image.docx file
///
/// # Returns
/// The extracted `ImageTemplate`
pub fn extract(path: &Path) -> Result<ImageTemplate> {
    if !path.exists() {
        return Err(Error::Template(format!(
            "Image template file not found: {}",
            path.display()
        )));
    }

    // Open DOCX as ZIP
    let file = std::fs::File::open(path)
        .map_err(|e| Error::Template(format!("Failed to open image template: {}", e)))?;
    let mut archive = ZipArchive::new(file)
        .map_err(|e| Error::Template(format!("Failed to read image template as ZIP: {}", e)))?;

    // Read word/document.xml
    let mut document_xml = String::new();
    {
        let mut doc_file = archive
            .by_name("word/document.xml")
            .map_err(|e| Error::Template(format!("Failed to find document.xml: {}", e)))?;
        doc_file
            .read_to_string(&mut document_xml)
            .map_err(|e| Error::Template(format!("Failed to read document.xml: {}", e)))?;
    }

    extract_from_xml(&document_xml)
}

fn extract_from_xml(xml: &str) -> Result<ImageTemplate> {
    let mut template = ImageTemplate::default();

    // Find <w:drawing>
    let drawing_pos = xml.find("<w:drawing>");
    if let Some(pos) = drawing_pos {
        let drawing_xml = extract_element(xml, pos, "</w:drawing>")?;

        // Extract from <pic:spPr>
        if let Some(sp_pr_pos) = drawing_xml.find("<pic:spPr>") {
            let sp_pr_xml = extract_element(&drawing_xml, sp_pr_pos, "</pic:spPr>")?;

            // Border: <a:ln>
            template.border = extract_border(&sp_pr_xml);

            // Shadow: <a:effectLst> -> <a:outerShdw>
            template.shadow = extract_shadow(&sp_pr_xml);
        }

        // Extract effect extent
        template.effect_extent = extract_effect_extent(&drawing_xml);

        // Extract image paragraph alignment (w:jc from the paragraph containing the drawing)
        if let Some(draw_pos) = drawing_pos {
            // Find the start of the paragraph containing this drawing
            if let Some(p_start) = xml[..draw_pos].rfind("<w:p") {
                // Find the end of pPr section
                if let Some(ppr_end) = xml[p_start..].find("</w:pPr>") {
                    let ppr_xml = &xml[p_start..p_start + ppr_end];
                    // Look for w:jc
                    if let Some(jc_pos) = ppr_xml.find("<w:jc") {
                        if let Some(val) = extract_attribute(&ppr_xml[jc_pos..], "w:val=") {
                            template.alignment = val;
                        }
                    }
                }
            }
        }
    }

    // Find caption paragraph
    if let Some(caption_pos) = xml.find("{{image_caption_prefix}}") {
        if let Some(p_start) = xml[..caption_pos].rfind("<w:p") {
            if let Some(p_end) = xml[p_start..].find("</w:p>") {
                let p_xml = &xml[p_start..p_start + p_end + 6];

                template.caption_runs = extract_caption_runs(p_xml);
                template.caption = extract_caption_style_from_p(p_xml);

                if let Some(draw_pos) = drawing_pos {
                    if p_start < draw_pos {
                        template.caption.position = "top".to_string();
                    } else {
                        template.caption.position = "bottom".to_string();
                    }
                }
            }
        }
    }

    Ok(template)
}

fn extract_border(sp_pr_xml: &str) -> Option<ImageBorder> {
    if let Some(ln_pos) = sp_pr_xml.find("<a:ln") {
        if let Some(ln_end) = sp_pr_xml[ln_pos..].find("</a:ln>") {
            let ln_xml = &sp_pr_xml[ln_pos..ln_pos + ln_end + 7];
            let mut border = ImageBorder {
                fill_type: "solid".to_string(),
                ..Default::default()
            };

            if let Some(width) = extract_attribute(ln_xml, "w=") {
                border.width = width.parse().ok();
            }

            if let Some(solid_fill_pos) = ln_xml.find("<a:solidFill") {
                let fill_xml = &ln_xml[solid_fill_pos..];
                if let Some(scheme_clr_pos) = fill_xml.find("<a:schemeClr") {
                    if let Some(val) = extract_attribute(&fill_xml[scheme_clr_pos..], "val=") {
                        border.color = val;
                        border.is_scheme_color = true;
                    }
                } else if let Some(srgb_clr_pos) = fill_xml.find("<a:srgbClr") {
                    if let Some(val) = extract_attribute(&fill_xml[srgb_clr_pos..], "val=") {
                        border.color = format!("#{}", val);
                        border.is_scheme_color = false;
                    }
                }
            } else if ln_xml.contains("<a:noFill") {
                border.fill_type = "none".to_string();
            }

            return Some(border);
        } else if let Some(_ln_end) = sp_pr_xml[ln_pos..].find("/>") {
            // Self-closing <a:ln /> probably means no border
            return None;
        }
    }
    None
}

fn extract_shadow(sp_pr_xml: &str) -> Option<ImageShadow> {
    if let Some(shadow_pos) = sp_pr_xml.find("<a:outerShdw") {
        let fragment = &sp_pr_xml[shadow_pos..];
        let mut shadow = ImageShadow::default();

        if let Some(blur) = extract_attribute(fragment, "blurRad=") {
            shadow.blur_radius = blur.parse().unwrap_or(0);
        }
        if let Some(dist) = extract_attribute(fragment, "dist=") {
            shadow.distance = dist.parse().unwrap_or(0);
        }
        if let Some(dir) = extract_attribute(fragment, "dir=") {
            shadow.direction = dir.parse().unwrap_or(0);
        }
        if let Some(algn) = extract_attribute(fragment, "algn=") {
            shadow.alignment = algn;
        }

        if let Some(srgb_pos) = fragment.find("<a:srgbClr") {
            let srgb_fragment = &fragment[srgb_pos..];
            if let Some(val) = extract_attribute(srgb_fragment, "val=") {
                shadow.color = val;
            }
            if let Some(alpha_pos) = srgb_fragment.find("<a:alpha") {
                if let Some(val) = extract_attribute(&srgb_fragment[alpha_pos..], "val=") {
                    shadow.alpha = val.parse().unwrap_or(100000);
                }
            }
        }

        return Some(shadow);
    }
    None
}

fn extract_effect_extent(drawing_xml: &str) -> EffectExtent {
    let mut extent = EffectExtent::default();
    if let Some(pos) = drawing_xml.find("<wp:effectExtent") {
        let fragment = &drawing_xml[pos..];
        if let Some(l) = extract_attribute(fragment, "l=") {
            extent.left = l.parse().unwrap_or(0);
        }
        if let Some(t) = extract_attribute(fragment, "t=") {
            extent.top = t.parse().unwrap_or(0);
        }
        if let Some(r) = extract_attribute(fragment, "r=") {
            extent.right = r.parse().unwrap_or(0);
        }
        if let Some(b) = extract_attribute(fragment, "b=") {
            extent.bottom = b.parse().unwrap_or(0);
        }
    }
    extent
}

fn extract_caption_runs(p_xml: &str) -> Vec<CaptionRun> {
    let mut runs = Vec::new();
    let mut pos = 0;
    while let Some(r_start) = p_xml[pos..].find("<w:r") {
        let abs_r_start = pos + r_start;
        if let Some(r_end) = p_xml[abs_r_start..].find("</w:r>") {
            let r_xml = &p_xml[abs_r_start..abs_r_start + r_end + 6];
            let mut run = CaptionRun::default();

            if r_xml.contains("image_caption_prefix") {
                run.placeholder = "prefix".to_string();
            } else if r_xml.contains("image_number") {
                run.placeholder = "number".to_string();
            } else if r_xml.contains("image_caption_text") {
                run.placeholder = "text".to_string();
            }

            let (font, size, color, bold, italic) = extract_run_properties_local(r_xml);
            run.font_family = font;
            run.font_size = size;
            run.font_color = color.replace("#", "");
            run.bold = bold;
            run.italic = italic;

            runs.push(run);
            pos = abs_r_start + r_end + 6;
        } else {
            break;
        }
    }
    runs
}

fn extract_caption_style_from_p(p_xml: &str) -> ImageCaptionStyle {
    let mut style = ImageCaptionStyle::default();

    let (font, size, color, bold, italic) = extract_run_properties_local(p_xml);
    style.font_family = font;
    style.font_size = size;
    style.font_color = color;
    style.bold = bold;
    style.italic = italic;

    if let Some(jc) = extract_attribute(p_xml, "w:jc w:val=") {
        style.alignment = jc;
    }

    if let Some(spacing_pos) = p_xml.find("<w:spacing") {
        if let Some(before) = extract_attribute(&p_xml[spacing_pos..], "w:before=") {
            if let Ok(v) = before.parse::<u32>() {
                style.spacing_before = v;
            }
        }
        if let Some(after) = extract_attribute(&p_xml[spacing_pos..], "w:after=") {
            if let Ok(v) = after.parse::<u32>() {
                style.spacing_after = v;
            }
        }
    }

    style
}

fn extract_element(xml: &str, start_pos: usize, close_tag: &str) -> Result<String> {
    let fragment = &xml[start_pos..];
    if let Some(end_pos) = fragment.find(close_tag) {
        Ok(fragment[..end_pos + close_tag.len()].to_string())
    } else {
        Err(Error::Template(format!(
            "Failed to find closing tag {}",
            close_tag
        )))
    }
}

fn extract_run_properties_local(xml: &str) -> (String, u32, String, bool, bool) {
    // Use the shared function with default defaults
    if let Some(rpr_start) = xml.find("<w:rPr") {
        if let Some(rpr_end) = xml[rpr_start..].find("</w:rPr>") {
            let rpr_xml = &xml[rpr_start..rpr_start + rpr_end + 8];
            let props = extract_run_properties(rpr_xml, &RunPropertiesDefaults::default());
            return (
                props.font_family,
                props.font_size,
                props.font_color,
                props.bold,
                props.italic,
            );
        }
    }
    // Return defaults if no rPr found
    let defaults = RunPropertiesDefaults::default();
    (
        defaults.font_family.to_string(),
        defaults.font_size,
        defaults.font_color.to_string(),
        false,
        false,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_image_template_default() {
        let template = ImageTemplate::default();

        assert_eq!(template.caption.position, "bottom");
        assert_eq!(template.caption.prefix, "Figure");
        assert!(template.caption.italic);
        assert_eq!(template.caption.alignment, "center");
        assert_eq!(template.default_width_percent, 80);
        assert_eq!(template.alignment, "center");
    }

    #[test]
    fn test_image_template_builder() {
        let template = ImageTemplate::new()
            .with_caption_prefix("รูปที่")
            .with_caption_position("top")
            .with_default_width(100);

        assert_eq!(template.caption.prefix, "รูปที่");
        assert_eq!(template.caption.position, "top");
        assert_eq!(template.default_width_percent, 100);
    }

    #[test]
    fn test_format_caption() {
        let template = ImageTemplate::default();
        let caption = template.format_caption("1.2", "System Architecture");

        assert_eq!(caption, "Figure 1.2: System Architecture");
    }

    #[test]
    fn test_extract_file_not_found() {
        let result = extract(Path::new("/nonexistent/image.docx"));
        assert!(result.is_err());
    }

    #[test]
    fn test_debug_extract_real_template() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("docs/template/image.docx");
        if !path.exists() {
            println!("Template file not found, skipping test");
            return;
        }

        let result = extract(&path);
        match result {
            Ok(template) => {
                println!("\n=== IMAGE TEMPLATE EXTRACTION DEBUG ===\n");
                println!("BORDER: {:?}", template.border);
                println!("SHADOW: {:?}", template.shadow);
                println!("EFFECT EXTENT: {:?}", template.effect_extent);
                println!("CAPTION POSITION: {}", template.caption.position);
                println!("CAPTION PREFIX: {}", template.caption.prefix);
                println!("CAPTION RUNS: {:?}", template.caption_runs);
                println!("\n=== END DEBUG ===\n");
            }
            Err(e) => {
                println!("Error extracting template: {:?}", e);
            }
        }
    }
}
