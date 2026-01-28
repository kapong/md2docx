//! Mermaid configuration

use serde::{Deserialize, Serialize};

/// Configuration for Mermaid diagram rendering
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MermaidConfig {
    /// Enable mermaid rendering
    #[serde(default = "default_enabled")]
    pub enabled: bool,

    /// Theme: default, forest, dark, neutral
    #[serde(default = "default_theme")]
    pub theme: String,

    /// Background color
    #[serde(default = "default_background")]
    pub background: String,

    /// Max width in pixels
    #[serde(default = "default_width")]
    pub width: u32,

    /// Scale factor for high DPI (2 = 2x)
    #[serde(default = "default_scale")]
    pub scale: u32,

    /// Enable caching
    #[serde(default = "default_cache")]
    pub cache: bool,
}

impl Default for MermaidConfig {
    fn default() -> Self {
        Self {
            enabled: default_enabled(),
            theme: default_theme(),
            background: default_background(),
            width: default_width(),
            scale: default_scale(),
            cache: default_cache(),
        }
    }
}

// Default value functions
fn default_enabled() -> bool {
    true
}
fn default_theme() -> String {
    "default".to_string()
}
fn default_background() -> String {
    "white".to_string()
}
fn default_width() -> u32 {
    800
}
fn default_scale() -> u32 {
    2
}
fn default_cache() -> bool {
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_config() {
        let config = MermaidConfig::default();
        assert!(config.enabled);
        assert_eq!(config.theme, "default");
        assert_eq!(config.width, 800);
    }
}
