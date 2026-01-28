//! Include directive resolution
//!
//! Resolves {!include:...} and {!code:...} directives by loading
//! external files and converting them to markdown blocks.

use std::fs;
use std::path::{Path, PathBuf};

use crate::error::{Error, Result};
use crate::parser::{parse_markdown, Block};

/// Configuration for include resolution
#[derive(Debug, Clone)]
pub struct IncludeConfig {
    /// Base directory for relative paths (usually the document directory)
    pub base_path: PathBuf,
    /// Root directory for code includes (from config: source_root)
    pub source_root: PathBuf,
    /// Maximum nesting depth to prevent infinite recursion
    pub max_depth: u32,
}

impl Default for IncludeConfig {
    fn default() -> Self {
        Self {
            base_path: PathBuf::from("."),
            source_root: PathBuf::from("."),
            max_depth: 10,
        }
    }
}

/// Resolver for include directives
pub struct IncludeResolver {
    config: IncludeConfig,
    /// Track included files to detect cycles
    include_stack: Vec<PathBuf>,
}

impl IncludeResolver {
    pub fn new(config: IncludeConfig) -> Self {
        Self {
            config,
            include_stack: Vec::new(),
        }
    }

    /// Resolve all include directives in a list of blocks
    /// Returns new blocks with includes expanded
    pub fn resolve_blocks(&mut self, blocks: Vec<Block>) -> Result<Vec<Block>> {
        let mut result = Vec::new();

        for block in blocks {
            match block {
                Block::Include { path, .. } => {
                    let included = self.resolve_include(&path)?;
                    result.extend(included);
                }
                Block::CodeInclude {
                    path,
                    start_line,
                    end_line,
                    lang,
                } => {
                    let code_block =
                        self.resolve_code(&path, start_line, end_line, lang.as_deref())?;
                    result.push(code_block);
                }
                Block::BlockQuote(inner) => {
                    let resolved_inner = self.resolve_blocks(inner)?;
                    result.push(Block::BlockQuote(resolved_inner));
                }
                Block::List {
                    ordered,
                    start,
                    items,
                } => {
                    let resolved_items = items
                        .into_iter()
                        .map(|item| {
                            let resolved_content = self.resolve_blocks(item.content)?;
                            Ok(crate::parser::ListItem {
                                content: resolved_content,
                                checked: item.checked,
                            })
                        })
                        .collect::<Result<Vec<_>>>()?;
                    result.push(Block::List {
                        ordered,
                        start,
                        items: resolved_items,
                    });
                }
                // Other blocks pass through unchanged
                other => result.push(other),
            }
        }

        Ok(result)
    }

    /// Resolve a markdown include directive
    fn resolve_include(&mut self, path: &str) -> Result<Vec<Block>> {
        let full_path = self.config.base_path.join(path);
        let canonical = full_path
            .canonicalize()
            .map_err(|e| Error::Include(format!("Cannot resolve path {}: {}", path, e)))?;

        // Check for cycles
        if self.include_stack.contains(&canonical) {
            return Err(Error::Include(format!(
                "Circular include detected: {} is already in the include stack",
                path
            )));
        }

        // Check depth limit
        if self.include_stack.len() >= self.config.max_depth as usize {
            return Err(Error::Include(format!(
                "Include depth exceeded (max {}): {}",
                self.config.max_depth, path
            )));
        }

        // Read the file
        let content = fs::read_to_string(&canonical)
            .map_err(|e| Error::Include(format!("Cannot read {}: {}", path, e)))?;

        // Push to stack before parsing (to detect cycles in nested includes)
        self.include_stack.push(canonical.clone());

        // Parse the included markdown
        let parsed = parse_markdown(&content);

        // Recursively resolve any nested includes
        let resolved = self.resolve_blocks(parsed.blocks)?;

        // Pop from stack
        self.include_stack.pop();

        Ok(resolved)
    }

    /// Resolve a code include directive
    fn resolve_code(
        &self,
        path: &str,
        start_line: Option<u32>,
        end_line: Option<u32>,
        lang_override: Option<&str>,
    ) -> Result<Block> {
        let full_path = self.config.source_root.join(path);

        let content = fs::read_to_string(&full_path)
            .map_err(|e| Error::Include(format!("Cannot read code file {}: {}", path, e)))?;

        // Extract lines if specified
        let lines: Vec<&str> = content.lines().collect();
        let start_idx = start_line
            .map(|n| (n.saturating_sub(1)) as usize)
            .unwrap_or(0);
        let end_idx = end_line.map(|n| n as usize).unwrap_or(lines.len());

        let extracted: String = lines
            .get(start_idx..end_idx.min(lines.len()))
            .unwrap_or(&[])
            .join("\n");

        // Infer language from extension if not specified
        let language = lang_override.map(String::from).or_else(|| {
            Path::new(path)
                .extension()
                .and_then(|e| e.to_str())
                .map(|ext| match ext {
                    "rs" => "rust",
                    "py" => "python",
                    "js" => "javascript",
                    "ts" => "typescript",
                    "go" => "go",
                    "rb" => "ruby",
                    "java" => "java",
                    "c" | "h" => "c",
                    "cpp" | "hpp" | "cc" => "cpp",
                    "sh" | "bash" => "bash",
                    "yaml" | "yml" => "yaml",
                    "json" => "json",
                    "toml" => "toml",
                    "xml" => "xml",
                    "html" => "html",
                    "css" => "css",
                    "sql" => "sql",
                    "md" => "markdown",
                    other => other,
                })
                .map(String::from)
        });

        Ok(Block::CodeBlock {
            lang: language,
            content: extracted,
            filename: Some(path.to_string()),
            highlight_lines: vec![],
            show_line_numbers: false,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use tempfile::TempDir;

    fn create_temp_file(dir: &TempDir, name: &str, content: &str) -> PathBuf {
        let path = dir.path().join(name);
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).unwrap();
        }
        let mut file = fs::File::create(&path).unwrap();
        file.write_all(content.as_bytes()).unwrap();
        path
    }

    #[test]
    fn test_resolve_code_include() {
        let temp_dir = TempDir::new().unwrap();
        create_temp_file(
            &temp_dir,
            "main.rs",
            "fn main() {\n    println!(\"Hello\");\n}\n",
        );

        let config = IncludeConfig {
            base_path: temp_dir.path().to_path_buf(),
            source_root: temp_dir.path().to_path_buf(),
            max_depth: 10,
        };

        let resolver = IncludeResolver::new(config);
        let result = resolver.resolve_code("main.rs", None, None, None).unwrap();

        match result {
            Block::CodeBlock {
                lang,
                content,
                filename,
                ..
            } => {
                assert_eq!(lang, Some("rust".to_string()));
                assert!(content.contains("fn main()"));
                assert_eq!(filename, Some("main.rs".to_string()));
            }
            _ => panic!("Expected CodeBlock"),
        }
    }

    #[test]
    fn test_resolve_code_with_line_range() {
        let temp_dir = TempDir::new().unwrap();
        create_temp_file(
            &temp_dir,
            "lines.txt",
            "line 1\nline 2\nline 3\nline 4\nline 5\n",
        );

        let config = IncludeConfig {
            base_path: temp_dir.path().to_path_buf(),
            source_root: temp_dir.path().to_path_buf(),
            max_depth: 10,
        };

        let resolver = IncludeResolver::new(config);
        let result = resolver
            .resolve_code("lines.txt", Some(2), Some(4), None)
            .unwrap();

        match result {
            Block::CodeBlock { content, .. } => {
                assert_eq!(content, "line 2\nline 3\nline 4");
            }
            _ => panic!("Expected CodeBlock"),
        }
    }

    #[test]
    fn test_include_config_default() {
        let config = IncludeConfig::default();
        assert_eq!(config.max_depth, 10);
    }

    #[test]
    fn test_language_inference() {
        let temp_dir = TempDir::new().unwrap();

        // Test various extensions
        for (file, expected_lang) in &[
            ("test.py", "python"),
            ("test.js", "javascript"),
            ("test.go", "go"),
            ("test.yaml", "yaml"),
        ] {
            create_temp_file(&temp_dir, file, "content");

            let config = IncludeConfig {
                source_root: temp_dir.path().to_path_buf(),
                ..Default::default()
            };

            let resolver = IncludeResolver::new(config);
            let result = resolver.resolve_code(file, None, None, None).unwrap();

            match result {
                Block::CodeBlock { lang, .. } => {
                    assert_eq!(lang, Some(expected_lang.to_string()), "Failed for {}", file);
                }
                _ => panic!("Expected CodeBlock"),
            }
        }
    }
}
