//! Image utilities for reading dimensions and calculating sizes

/// Image dimensions in pixels
#[derive(Debug, Clone, Copy)]
pub struct ImageDimensions {
    pub width: u32,
    pub height: u32,
}

impl ImageDimensions {
    /// Calculate aspect ratio (width / height)
    pub fn aspect_ratio(&self) -> f64 {
        if self.height == 0 {
            1.0
        } else {
            self.width as f64 / self.height as f64
        }
    }
}

/// Read image dimensions from image data
/// Supports PNG, JPEG, GIF, BMP, and SVG
pub fn read_image_dimensions(data: &[u8]) -> Option<ImageDimensions> {
    // Try PNG
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        return read_png_dimensions(data);
    }

    // Try JPEG
    if data.starts_with(b"\xFF\xD8\xFF") {
        return read_jpeg_dimensions(data);
    }

    // Try GIF
    if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        return read_gif_dimensions(data);
    }

    // Try BMP
    if data.starts_with(b"BM") {
        return read_bmp_dimensions(data);
    }

    // Try SVG (XML-based)
    if data.starts_with(b"<") || data.starts_with(b"<?xml") {
        return read_svg_dimensions(data);
    }

    None
}

fn read_png_dimensions(data: &[u8]) -> Option<ImageDimensions> {
    // PNG dimensions are at bytes 16-24
    // Width: bytes 16-19 (big-endian)
    // Height: bytes 20-23 (big-endian)
    if data.len() >= 24 {
        let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
        let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
        Some(ImageDimensions { width, height })
    } else {
        None
    }
}

fn read_jpeg_dimensions(data: &[u8]) -> Option<ImageDimensions> {
    // JPEG is more complex - scan for SOF markers
    let mut i = 2;
    while i < data.len() - 1 {
        if data[i] == 0xFF {
            let marker = data[i + 1];
            // SOF0, SOF1, SOF2 markers contain dimensions
            if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2) && i + 9 < data.len() {
                let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
                let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
                return Some(ImageDimensions { width, height });
            }
            // Skip marker segment
            if marker != 0x00
                && marker != 0x01
                && !(0xD0..=0xD9).contains(&marker)
                && i + 3 < data.len()
            {
                let len = u16::from_be_bytes([data[i + 2], data[i + 3]]) as usize;
                i += len + 2;
                continue;
            }
        }
        i += 1;
    }
    None
}

fn read_gif_dimensions(data: &[u8]) -> Option<ImageDimensions> {
    // GIF dimensions at bytes 6-10 (little-endian)
    if data.len() >= 10 {
        let width = u16::from_le_bytes([data[6], data[7]]) as u32;
        let height = u16::from_le_bytes([data[8], data[9]]) as u32;
        Some(ImageDimensions { width, height })
    } else {
        None
    }
}

fn read_bmp_dimensions(data: &[u8]) -> Option<ImageDimensions> {
    // BMP dimensions at bytes 18-26
    if data.len() >= 26 {
        let width = i32::from_le_bytes([data[18], data[19], data[20], data[21]]) as u32;
        let height = i32::from_le_bytes([data[22], data[23], data[24], data[25]]);
        Some(ImageDimensions {
            width,
            height: height.unsigned_abs(),
        })
    } else {
        None
    }
}

fn read_svg_dimensions(data: &[u8]) -> Option<ImageDimensions> {
    // Parse SVG width/height attributes or viewBox
    let data_str = String::from_utf8_lossy(data);

    // Try to extract width and height attributes
    let width_re = regex::Regex::new(r#"width="([^"]+)""#).ok()?;
    let height_re = regex::Regex::new(r#"height="([^"]+)""#).ok()?;

    let width_caps = width_re.captures(&data_str)?;
    let height_caps = height_re.captures(&data_str)?;

    let width_str = width_caps.get(1)?.as_str();
    let height_str = height_caps.get(1)?.as_str();

    // Parse dimension values (may include units like "px", "pt", etc.)
    let width = parse_svg_dimension(width_str)?;
    let height = parse_svg_dimension(height_str)?;

    Some(ImageDimensions { width, height })
}

fn parse_svg_dimension(s: &str) -> Option<u32> {
    // Extract numeric part from strings like "100px", "100", "100.5"
    let num_str: String = s
        .chars()
        .take_while(|c| c.is_ascii_digit() || *c == '.')
        .collect();
    num_str.parse::<f64>().ok().map(|n| n as u32)
}

/// Calculate image size in EMUs for DOCX
///
/// # Arguments
/// * `dims` - Image dimensions in pixels
/// * `target_dpi` - Target DPI (default 96)
/// * `max_width_inches` - Maximum width in inches (default 6.0 for A4 with margins)
/// * `max_height_inches` - Maximum height in inches (default 9.0)
///
/// # Returns
/// (width_emu, height_emu) sized to fit within constraints while preserving aspect ratio
pub fn calculate_image_size_emu(
    dims: ImageDimensions,
    target_dpi: f64,
    max_width_inches: f64,
    max_height_inches: f64,
) -> (i64, i64) {
    const EMU_PER_INCH: f64 = 914400.0;

    // Calculate original size in inches at target DPI
    let width_inches = dims.width as f64 / target_dpi;
    let height_inches = dims.height as f64 / target_dpi;

    // Check if we need to scale down

    let scale_w = max_width_inches / width_inches;
    let scale_h = max_height_inches / height_inches;
    let scale = scale_w.min(scale_h).min(1.0); // Never scale up, only down

    // Apply scaling
    let final_width_inches = width_inches * scale;
    let final_height_inches = height_inches * scale;

    // Convert to EMUs
    let width_emu = (final_width_inches * EMU_PER_INCH) as i64;
    let height_emu = (final_height_inches * EMU_PER_INCH) as i64;

    (width_emu, height_emu)
}

/// Default image sizing with sensible defaults for DOCX
/// - 96 DPI (standard screen resolution)
/// - Max 6 inches width (A4 page with 1-inch margins)
/// - Max 9 inches height
pub fn default_image_size_emu(dims: ImageDimensions) -> (i64, i64) {
    calculate_image_size_emu(dims, 96.0, 6.0, 9.0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_png_dimensions() {
        // Create minimal PNG header
        let mut data = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        data.extend_from_slice(&[0x00, 0x00, 0x00, 0x0D]); // IHDR length
        data.extend_from_slice(b"IHDR");
        data.extend_from_slice(&[0x00, 0x00, 0x01, 0x00]); // Width: 256
        data.extend_from_slice(&[0x00, 0x00, 0x00, 0x80]); // Height: 128

        let dims = read_png_dimensions(&data).unwrap();
        assert_eq!(dims.width, 256);
        assert_eq!(dims.height, 128);
    }

    #[test]
    fn test_calculate_size() {
        let dims = ImageDimensions {
            width: 1920,
            height: 1080,
        };
        let (w, h) = default_image_size_emu(dims);

        // At 96 DPI, 1920px = 20 inches, but max is 6 inches
        // So it should be scaled to 6 inches wide
        // Height should maintain aspect ratio: 6 * (1080/1920) = 3.375 inches
        assert!(w > 0);
        assert!(h > 0);

        // Verify aspect ratio is preserved
        let aspect = w as f64 / h as f64;
        assert!((aspect - 1920.0 / 1080.0).abs() < 0.01);
    }
}
