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
///
/// Content inside fenced code blocks is preserved unmodified.
pub fn resolve_image_paths(content: &str, file_path: &Path) -> String {
    if let Some(parent) = file_path.parent() {
        let image_regex = Regex::new(r"!\[(.*?)\]\s*\((.*?)\)").expect("Invalid regex");

        // Split content into code-block and non-code-block regions,
        // only replacing image paths outside code blocks.
        let mut result = String::with_capacity(content.len());
        let mut remaining = content;

        while !remaining.is_empty() {
            // Find the next fenced code block opening
            if let Some(fence_start) = find_code_fence_start(remaining) {
                // Process text before the code block
                let before = &remaining[..fence_start.offset];
                result.push_str(&replace_image_paths_in_text(before, parent, &image_regex));

                // Find the matching closing fence
                let fence_content_start = fence_start.offset;
                if let Some(fence_end) = find_code_fence_end(
                    &remaining[fence_content_start..],
                    fence_start.backtick_count,
                ) {
                    // Append the entire code block verbatim
                    let block_end = fence_content_start + fence_end;
                    result.push_str(&remaining[fence_content_start..block_end]);
                    remaining = &remaining[block_end..];
                } else {
                    // No closing fence found; treat the rest as a code block (verbatim)
                    result.push_str(&remaining[fence_content_start..]);
                    remaining = "";
                }
            } else {
                // No more code blocks, process the rest
                result.push_str(&replace_image_paths_in_text(remaining, parent, &image_regex));
                remaining = "";
            }
        }

        result
    } else {
        content.to_string()
    }
}

/// Information about a fenced code block opening
struct CodeFenceStart {
    offset: usize,
    backtick_count: usize,
}

/// Find the start of the next fenced code block (``` or ~~~) at the beginning of a line
fn find_code_fence_start(text: &str) -> Option<CodeFenceStart> {
    for (i, line) in text.lines().enumerate() {
        let trimmed = line.trim_start();
        let fence_char = if trimmed.starts_with("```") {
            Some('`')
        } else if trimmed.starts_with("~~~") {
            Some('~')
        } else {
            None
        };
        if let Some(ch) = fence_char {
            let count = trimmed.chars().take_while(|&c| c == ch).count();
            if count >= 3 {
                // Calculate byte offset of this line in the text
                let offset: usize = text.lines().take(i).map(|l| l.len() + 1).sum();
                // Clamp to text length (last line may not have trailing newline)
                let offset = offset.min(text.len());
                return Some(CodeFenceStart {
                    offset,
                    backtick_count: count,
                });
            }
        }
    }
    None
}

/// Find the end of a fenced code block (matching closing fence)
/// Returns byte offset past the closing fence line (including its newline)
fn find_code_fence_end(text: &str, opening_count: usize) -> Option<usize> {
    let fence_char = text.trim_start().chars().next().unwrap_or('`');
    let mut offset = 0;
    let mut first_line = true;
    for line in text.lines() {
        offset += line.len() + 1; // +1 for newline
        if first_line {
            first_line = false;
            continue; // Skip the opening fence line
        }
        let trimmed = line.trim_start();
        let count = trimmed.chars().take_while(|&c| c == fence_char).count();
        // Closing fence: same or more fence chars, no info string (only whitespace after)
        if count >= opening_count {
            let after_fence = &trimmed[count..];
            if after_fence.trim().is_empty() {
                return Some(offset.min(text.len()));
            }
        }
    }
    None
}

/// Replace image paths in a text segment (outside code blocks)
fn replace_image_paths_in_text(text: &str, parent: &Path, image_regex: &Regex) -> String {
    image_regex
        .replace_all(text, |caps: &regex::Captures| {
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
