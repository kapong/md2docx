//! Markdown file processing utilities

use regex::Regex;
use std::path::Path;

/// Strip YAML frontmatter from markdown content
///
/// Frontmatter is delimited by `---` at the start of the file.
///
/// # Example
/// ```
/// use md2docx::project::strip_frontmatter;
///
/// let content = "---\ntitle: Test\n---\n\n# Heading";
/// let stripped = strip_frontmatter(content);
/// assert!(stripped.contains("# Heading"));
/// assert!(stripped.starts_with("\n"));
/// ```
pub fn strip_frontmatter(content: &str) -> String {
    if !content.starts_with("---") {
        return content.to_string();
    }

    let lines: Vec<&str> = content.lines().collect();
    let mut closing_line = None;

    for (i, line) in lines.iter().enumerate().skip(1) {
        if line.trim() == "---" {
            closing_line = Some(i);
            break;
        }
    }

    match closing_line {
        Some(idx) => lines[idx + 1..].join("\n"),
        None => content.to_string(),
    }
}

/// Rewrite relative image paths in markdown content to be relative to the markdown file's directory
///
/// This ensures that when multiple markdown files are combined, their relative image
/// paths still resolve correctly.
pub fn resolve_image_paths(content: &str, file_path: &Path) -> String {
    if let Some(parent) = file_path.parent() {
        let image_regex = Regex::new(r"!\[(.*?)\]\s*\((.*?)\)").expect("Invalid regex");

        image_regex
            .replace_all(content, |caps: &regex::Captures| {
                let alt = &caps[1];
                let raw_link = &caps[2];
                let link_content = raw_link.trim();

                // Split url and optional title
                let (url, title_suffix) = match link_content.find(char::is_whitespace) {
                    Some(idx) => (&link_content[..idx], &link_content[idx..]),
                    None => (link_content, ""),
                };

                // Skip absolute URLs, absolute paths, or data URIs
                if url.starts_with("http://")
                    || url.starts_with("https://")
                    || url.starts_with("/")
                    || url.starts_with("data:")
                    || Path::new(url).is_absolute()
                {
                    return caps[0].to_string();
                }

                // Resolve relative to file parent
                let new_path = parent.join(url);
                let new_path_str = new_path.to_string_lossy().replace('\\', "/");

                format!("![{}]({}{})", alt, new_path_str, title_suffix)
            })
            .to_string()
    } else {
        content.to_string()
    }
}

/// Extract content from cover.md for the `{{inside}}` placeholder
///
/// Returns the content after YAML frontmatter (if any), with image paths
/// fixed to be relative to the project root.
pub fn extract_cover_inside_content(base_dir: &Path) -> Option<String> {
    let cover_path = base_dir.join("cover.md");
    if !cover_path.exists() {
        return None;
    }

    let content = std::fs::read_to_string(&cover_path).ok()?;
    let inside = strip_frontmatter(&content);
    let trimmed = inside.trim();

    if trimmed.is_empty() {
        return None;
    }

    // Fix image paths to be relative to project root
    let fixed_content = if base_dir.components().count() > 0 {
        let prefix = format!("{}/", base_dir.display());
        trimmed.replace("assets/", &format!("{}assets/", prefix))
    } else {
        trimmed.to_string()
    };

    Some(fixed_content)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_strip_frontmatter_with_frontmatter() {
        let content = "---\ntitle: Test\nauthor: Me\n---\n\n# Heading\n\nContent";
        let result = strip_frontmatter(content);
        assert!(result.starts_with("\n# Heading"));
    }

    #[test]
    fn test_strip_frontmatter_without_frontmatter() {
        let content = "# Heading\n\nContent";
        let result = strip_frontmatter(content);
        assert_eq!(result, content);
    }

    #[test]
    fn test_strip_frontmatter_unclosed() {
        let content = "---\ntitle: Test\n# Heading";
        let result = strip_frontmatter(content);
        assert_eq!(result, content);
    }

    #[test]
    fn test_resolve_image_paths_relative() {
        let content = "![Image](img.png)";
        let file_path = Path::new("docs/chapter1.md");
        let result = resolve_image_paths(content, file_path);
        assert_eq!(result, "![Image](docs/img.png)");
    }

    #[test]
    fn test_resolve_image_paths_absolute_url() {
        let content = "![Image](https://example.com/img.png)";
        let file_path = Path::new("docs/chapter1.md");
        let result = resolve_image_paths(content, file_path);
        assert_eq!(result, content);
    }
}
