//! File discovery for md2docx projects
//!
//! This module handles discovering and organizing markdown files in a project directory,
//! including cover pages, chapters, appendices, and bibliography files.

use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use crate::config::ProjectConfig;
use crate::Result;

/// Discovered project structure
#[derive(Debug, Clone, Default)]
pub struct DiscoveredProject {
    /// Cover page file (cover.md)
    pub cover: Option<PathBuf>,
    /// Chapter files sorted by number (ch##_*.md)
    pub chapters: Vec<ChapterFile>,
    /// Appendix files sorted (ap##_*.md)
    pub appendices: Vec<AppendixFile>,
    /// Bibliography file (bibliography.md or references.md)
    pub bibliography: Option<PathBuf>,
    /// Config file location (md2docx.toml)
    pub config_file: Option<PathBuf>,
    /// Base directory
    pub base_dir: PathBuf,
}

/// A discovered chapter file
#[derive(Debug, Clone)]
pub struct ChapterFile {
    /// Chapter number (e.g., 1 for ch01_intro.md)
    pub number: u32,
    /// Full path to the file
    pub path: PathBuf,
    /// Extracted name (e.g., "intro" from ch01_intro.md)
    pub name: String,
}

/// A discovered appendix file
#[derive(Debug, Clone)]
pub struct AppendixFile {
    /// Appendix number (e.g., 1 for ap01_troubleshooting.md)
    pub number: u32,
    /// Full path to the file
    pub path: PathBuf,
    /// Extracted name (e.g., "troubleshooting" from ap01_troubleshooting.md)
    pub name: String,
    /// Letter label (A, B, C...)
    pub letter: char,
}

impl DiscoveredProject {
    /// Discover project files in a directory using default patterns
    ///
    /// # Arguments
    /// * `dir` - The directory to search for project files
    ///
    /// # Returns
    /// A `Result` containing the discovered project structure
    ///
    /// # Example
    /// ```rust,no_run
    /// use md2docx::discovery::DiscoveredProject;
    /// use std::path::Path;
    ///
    /// let project = DiscoveredProject::discover(Path::new("./my-docs"))?;
    /// println!("Found {} chapters", project.chapters.len());
    /// # Ok::<(), md2docx::Error>(())
    /// ```
    #[cfg(not(target_arch = "wasm32"))]
    pub fn discover(dir: &Path) -> Result<Self> {
        Self::discover_with_config(dir, &ProjectConfig::default())
    }

    /// Discover project files using custom config patterns
    ///
    /// # Arguments
    /// * `dir` - The directory to search for project files
    /// * `config` - Configuration with custom patterns for file discovery
    ///
    /// # Returns
    /// A `Result` containing the discovered project structure
    #[cfg(not(target_arch = "wasm32"))]
    pub fn discover_with_config(dir: &Path, config: &ProjectConfig) -> Result<Self> {
        let base_dir = dir.canonicalize()?;

        // Look for config file
        let config_file = base_dir.join("md2docx.toml");
        let config_file = if config_file.exists() {
            Some(config_file)
        } else {
            None
        };

        // Look for cover page (case-insensitive)
        let cover = Self::find_cover(&base_dir);

        // Find chapter files
        let chapters = Self::find_chapters(&base_dir, &config.chapters.pattern)?;

        // Find appendix files
        let appendices = Self::find_appendices(&base_dir, &config.appendices.pattern)?;

        // Look for bibliography (case-insensitive)
        let bibliography = Self::find_bibliography(&base_dir);

        Ok(DiscoveredProject {
            cover,
            chapters,
            appendices,
            bibliography,
            config_file,
            base_dir,
        })
    }

    /// Get all markdown files in order (cover, chapters, appendices, bibliography)
    ///
    /// # Returns
    /// A vector of references to all markdown file paths in document order
    pub fn all_files(&self) -> Vec<&PathBuf> {
        let mut files = Vec::new();
        if let Some(ref cover) = self.cover {
            files.push(cover);
        }
        for ch in &self.chapters {
            files.push(&ch.path);
        }
        for ap in &self.appendices {
            files.push(&ap.path);
        }
        if let Some(ref bib) = self.bibliography {
            files.push(bib);
        }
        files
    }

    /// Check if this looks like a valid project directory
    ///
    /// A valid project has at least one chapter or a cover page.
    ///
    /// # Returns
    /// `true` if the directory contains valid project files
    pub fn is_valid(&self) -> bool {
        !self.chapters.is_empty() || self.cover.is_some()
    }

    /// Find cover page file (case-insensitive)
    #[cfg(not(target_arch = "wasm32"))]
    fn find_cover(base_dir: &Path) -> Option<PathBuf> {
        let cover_names = ["cover.md", "COVER.md", "Cover.md"];
        for name in &cover_names {
            let path = base_dir.join(name);
            if path.exists() {
                return Some(path);
            }
        }
        None
    }

    /// Find chapter files matching pattern
    #[cfg(not(target_arch = "wasm32"))]
    fn find_chapters(base_dir: &Path, _pattern: &str) -> Result<Vec<ChapterFile>> {
        use glob::glob;

        let mut chapters = Vec::new();
        let pattern_str = format!("{}/*.md", base_dir.display());

        for entry in glob(&pattern_str).map_err(|e| {
            crate::Error::Io(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("Invalid glob pattern: {}", e),
            ))
        })? {
            match entry {
                Ok(path) => {
                    if let Some(filename) = path.file_name().and_then(|n| n.to_str()) {
                        if let Some((number, name)) = parse_chapter_filename(filename) {
                            chapters.push(ChapterFile { number, path, name });
                        }
                    }
                }
                Err(e) => {
                    // Log warning but continue
                    eprintln!("Warning: Error reading file: {}", e);
                }
            }
        }

        // Sort by number (not alphabetically)
        chapters.sort_by(|a, b| a.number.cmp(&b.number));

        Ok(chapters)
    }

    /// Find appendix files matching pattern
    #[cfg(not(target_arch = "wasm32"))]
    fn find_appendices(base_dir: &Path, _pattern: &str) -> Result<Vec<AppendixFile>> {
        use glob::glob;

        let mut appendices = Vec::new();
        let pattern_str = format!("{}/*.md", base_dir.display());

        for entry in glob(&pattern_str).map_err(|e| {
            crate::Error::Io(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("Invalid glob pattern: {}", e),
            ))
        })? {
            match entry {
                Ok(path) => {
                    if let Some(filename) = path.file_name().and_then(|n| n.to_str()) {
                        if let Some((number, name)) = parse_appendix_filename(filename) {
                            // Convert number to letter (1 -> A, 2 -> B, etc.)
                            let letter = if number > 0 && number <= 26 {
                                (b'A' + (number - 1) as u8) as char
                            } else {
                                // Fallback for numbers beyond Z
                                '?'
                            };

                            appendices.push(AppendixFile {
                                number,
                                path,
                                name,
                                letter,
                            });
                        }
                    }
                }
                Err(e) => {
                    eprintln!("Warning: Error reading file: {}", e);
                }
            }
        }

        // Sort by number
        appendices.sort_by(|a, b| a.number.cmp(&b.number));

        Ok(appendices)
    }

    /// Find bibliography file (case-insensitive)
    #[cfg(not(target_arch = "wasm32"))]
    fn find_bibliography(base_dir: &Path) -> Option<PathBuf> {
        let bib_names = [
            "bibliography.md",
            "BIBLIOGRAPHY.md",
            "Bibliography.md",
            "references.md",
            "REFERENCES.md",
            "References.md",
        ];
        for name in &bib_names {
            let path = base_dir.join(name);
            if path.exists() {
                return Some(path);
            }
        }
        None
    }
}

/// Parse chapter number and name from filename
///
/// Supports patterns like:
/// - `ch01_introduction.md` -> (1, "introduction")
/// - `ch02_setup.md` -> (2, "setup")
/// - `ch10_advanced.md` -> (10, "advanced")
///
/// # Arguments
/// * `filename` - The filename to parse (e.g., "ch01_introduction.md")
///
/// # Returns
/// `Some((number, name))` if the filename matches the chapter pattern, `None` otherwise
pub fn parse_chapter_filename(filename: &str) -> Option<(u32, String)> {
    // Remove .md extension
    let stem = filename.strip_suffix(".md")?;

    // Pattern: ch##_name or ch##_name
    let stem_lower = stem.to_lowercase();

    if !stem_lower.starts_with("ch") {
        return None;
    }

    // Find the underscore after the number
    let rest = &stem[2..]; // Skip "ch"
    let underscore_pos = rest.find('_')?;

    // Extract number part
    let number_str = &rest[..underscore_pos];
    let number: u32 = number_str.parse().ok()?;

    // Extract name part (after underscore)
    let name = rest[underscore_pos + 1..].to_string();

    Some((number, name))
}

/// Parse appendix number and name from filename
///
/// Supports patterns like:
/// - `ap01_troubleshooting.md` -> (1, "troubleshooting")
/// - `ap02_glossary.md` -> (2, "glossary")
///
/// # Arguments
/// * `filename` - The filename to parse (e.g., "ap01_troubleshooting.md")
///
/// # Returns
/// `Some((number, name))` if the filename matches the appendix pattern, `None` otherwise
pub fn parse_appendix_filename(filename: &str) -> Option<(u32, String)> {
    // Remove .md extension
    let stem = filename.strip_suffix(".md")?;

    // Pattern: ap##_name
    let stem_lower = stem.to_lowercase();

    if !stem_lower.starts_with("ap") {
        return None;
    }

    // Find the underscore after the number
    let rest = &stem[2..]; // Skip "ap"
    let underscore_pos = rest.find('_')?;

    // Extract number part
    let number_str = &rest[..underscore_pos];
    let number: u32 = number_str.parse().ok()?;

    // Extract name part (after underscore)
    let name = rest[underscore_pos + 1..].to_string();

    Some((number, name))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_chapter_filename_valid() {
        assert_eq!(
            parse_chapter_filename("ch01_introduction.md"),
            Some((1, "introduction".to_string()))
        );
        assert_eq!(
            parse_chapter_filename("ch02_setup.md"),
            Some((2, "setup".to_string()))
        );
        assert_eq!(
            parse_chapter_filename("ch10_advanced.md"),
            Some((10, "advanced".to_string()))
        );
        assert_eq!(
            parse_chapter_filename("ch99_final.md"),
            Some((99, "final".to_string()))
        );
    }

    #[test]
    fn test_parse_chapter_filename_case_insensitive() {
        assert_eq!(
            parse_chapter_filename("CH01_INTRODUCTION.md"),
            Some((1, "INTRODUCTION".to_string()))
        );
        assert_eq!(
            parse_chapter_filename("Ch01_Introduction.md"),
            Some((1, "Introduction".to_string()))
        );
    }

    #[test]
    fn test_parse_chapter_filename_invalid() {
        assert_eq!(parse_chapter_filename("introduction.md"), None);
        assert_eq!(parse_chapter_filename("chapter01.md"), None);
        assert_eq!(parse_chapter_filename("ch01.md"), None); // No underscore
        assert_eq!(parse_chapter_filename("ch01_introduction.txt"), None); // Not .md
        assert_eq!(parse_chapter_filename("chxx_introduction.md"), None); // Invalid number
    }

    #[test]
    fn test_parse_appendix_filename_valid() {
        assert_eq!(
            parse_appendix_filename("ap01_troubleshooting.md"),
            Some((1, "troubleshooting".to_string()))
        );
        assert_eq!(
            parse_appendix_filename("ap02_glossary.md"),
            Some((2, "glossary".to_string()))
        );
        assert_eq!(
            parse_appendix_filename("ap05_faq.md"),
            Some((5, "faq".to_string()))
        );
    }

    #[test]
    fn test_parse_appendix_filename_case_insensitive() {
        assert_eq!(
            parse_appendix_filename("AP01_TROUBLESHOOTING.md"),
            Some((1, "TROUBLESHOOTING".to_string()))
        );
        assert_eq!(
            parse_appendix_filename("Ap01_Troubleshooting.md"),
            Some((1, "Troubleshooting".to_string()))
        );
    }

    #[test]
    fn test_parse_appendix_filename_invalid() {
        assert_eq!(parse_appendix_filename("troubleshooting.md"), None);
        assert_eq!(parse_appendix_filename("appendix01.md"), None);
        assert_eq!(parse_appendix_filename("ap01.md"), None); // No underscore
        assert_eq!(parse_appendix_filename("ap01_troubleshooting.txt"), None); // Not .md
        assert_eq!(parse_appendix_filename("apxx_troubleshooting.md"), None); // Invalid number
    }

    #[test]
    fn test_discovered_project_default() {
        let project = DiscoveredProject::default();
        assert!(project.cover.is_none());
        assert!(project.chapters.is_empty());
        assert!(project.appendices.is_empty());
        assert!(project.bibliography.is_none());
        assert!(project.config_file.is_none());
        assert!(!project.is_valid());
    }

    #[test]
    fn test_discovered_project_all_files() {
        let mut project = DiscoveredProject::default();
        project.base_dir = PathBuf::from("/test");

        // Add some files
        project.cover = Some(PathBuf::from("/test/cover.md"));
        project.chapters.push(ChapterFile {
            number: 1,
            path: PathBuf::from("/test/ch01_intro.md"),
            name: "intro".to_string(),
        });
        project.chapters.push(ChapterFile {
            number: 2,
            path: PathBuf::from("/test/ch02_setup.md"),
            name: "setup".to_string(),
        });
        project.appendices.push(AppendixFile {
            number: 1,
            path: PathBuf::from("/test/ap01_troubleshooting.md"),
            name: "troubleshooting".to_string(),
            letter: 'A',
        });
        project.bibliography = Some(PathBuf::from("/test/bibliography.md"));

        let files = project.all_files();
        assert_eq!(files.len(), 5);
        assert_eq!(files[0], &PathBuf::from("/test/cover.md"));
        assert_eq!(files[1], &PathBuf::from("/test/ch01_intro.md"));
        assert_eq!(files[2], &PathBuf::from("/test/ch02_setup.md"));
        assert_eq!(files[3], &PathBuf::from("/test/ap01_troubleshooting.md"));
        assert_eq!(files[4], &PathBuf::from("/test/bibliography.md"));
    }

    #[test]
    fn test_discovered_project_is_valid() {
        let mut project = DiscoveredProject::default();

        // Empty project is not valid
        assert!(!project.is_valid());

        // Project with cover is valid
        project.cover = Some(PathBuf::from("/test/cover.md"));
        assert!(project.is_valid());

        // Project with chapters is valid
        project.cover = None;
        project.chapters.push(ChapterFile {
            number: 1,
            path: PathBuf::from("/test/ch01_intro.md"),
            name: "intro".to_string(),
        });
        assert!(project.is_valid());
    }

    #[test]
    fn test_appendix_letter_assignment() {
        // Test that letters are assigned correctly (A, B, C...)
        let appendix1 = AppendixFile {
            number: 1,
            path: PathBuf::from("/test/ap01.md"),
            name: "test".to_string(),
            letter: 'A',
        };
        assert_eq!(appendix1.letter, 'A');

        let appendix2 = AppendixFile {
            number: 2,
            path: PathBuf::from("/test/ap02.md"),
            name: "test".to_string(),
            letter: 'B',
        };
        assert_eq!(appendix2.letter, 'B');

        let appendix26 = AppendixFile {
            number: 26,
            path: PathBuf::from("/test/ap26.md"),
            name: "test".to_string(),
            letter: 'Z',
        };
        assert_eq!(appendix26.letter, 'Z');
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_discover_project_directory() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let test_dir = temp_dir.join("md2docx_test_discovery");
        fs::create_dir_all(&test_dir).unwrap();

        // Create test files
        fs::write(test_dir.join("cover.md"), "# Cover").unwrap();
        fs::write(test_dir.join("ch01_introduction.md"), "# Chapter 1").unwrap();
        fs::write(test_dir.join("ch02_setup.md"), "# Chapter 2").unwrap();
        fs::write(test_dir.join("ch10_advanced.md"), "# Chapter 10").unwrap();
        fs::write(test_dir.join("ap01_troubleshooting.md"), "# Appendix A").unwrap();
        fs::write(test_dir.join("bibliography.md"), "# References").unwrap();

        // Discover project
        let project = DiscoveredProject::discover(&test_dir).unwrap();

        // Verify results
        assert!(project.is_valid());
        assert!(project.cover.is_some());
        assert_eq!(project.chapters.len(), 3);
        assert_eq!(project.appendices.len(), 1);
        assert!(project.bibliography.is_some());

        // Verify chapter order (numeric, not alphabetical)
        assert_eq!(project.chapters[0].number, 1);
        assert_eq!(project.chapters[1].number, 2);
        assert_eq!(project.chapters[2].number, 10);

        // Verify chapter names
        assert_eq!(project.chapters[0].name, "introduction");
        assert_eq!(project.chapters[1].name, "setup");
        assert_eq!(project.chapters[2].name, "advanced");

        // Verify appendix
        assert_eq!(project.appendices[0].number, 1);
        assert_eq!(project.appendices[0].name, "troubleshooting");
        assert_eq!(project.appendices[0].letter, 'A');

        // Cleanup
        fs::remove_dir_all(test_dir).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_discover_project_case_insensitive() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let test_dir = temp_dir.join("md2docx_test_case");
        fs::create_dir_all(&test_dir).unwrap();

        // Create files with different cases
        fs::write(test_dir.join("COVER.md"), "# Cover").unwrap();
        fs::write(test_dir.join("CH01_INTRODUCTION.md"), "# Chapter 1").unwrap();
        fs::write(test_dir.join("REFERENCES.md"), "# References").unwrap();

        let project = DiscoveredProject::discover(&test_dir).unwrap();

        assert!(project.cover.is_some());
        assert_eq!(project.chapters.len(), 1);
        assert!(project.bibliography.is_some());

        fs::remove_dir_all(test_dir).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_discover_project_empty_directory() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let test_dir = temp_dir.join("md2docx_test_empty");
        fs::create_dir_all(&test_dir).unwrap();

        let project = DiscoveredProject::discover(&test_dir).unwrap();

        assert!(!project.is_valid());
        assert!(project.cover.is_none());
        assert!(project.chapters.is_empty());
        assert!(project.appendices.is_empty());
        assert!(project.bibliography.is_none());

        fs::remove_dir_all(test_dir).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_discover_project_with_config_file() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let test_dir = temp_dir.join("md2docx_test_config");
        fs::create_dir_all(&test_dir).unwrap();

        // Create config file
        fs::write(test_dir.join("md2docx.toml"), "# Config").unwrap();
        fs::write(test_dir.join("ch01_test.md"), "# Chapter 1").unwrap();

        let project = DiscoveredProject::discover(&test_dir).unwrap();

        assert!(project.config_file.is_some());
        assert_eq!(project.chapters.len(), 1);

        fs::remove_dir_all(test_dir).unwrap();
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn test_discover_project_numeric_sorting() {
        use std::fs;

        let temp_dir = std::env::temp_dir();
        let test_dir = temp_dir.join("md2docx_test_sort");
        fs::create_dir_all(&test_dir).unwrap();

        // Create chapters that would sort incorrectly alphabetically
        fs::write(test_dir.join("ch2_second.md"), "# Chapter 2").unwrap();
        fs::write(test_dir.join("ch10_tenth.md"), "# Chapter 10").unwrap();
        fs::write(test_dir.join("ch1_first.md"), "# Chapter 1").unwrap();

        let project = DiscoveredProject::discover(&test_dir).unwrap();

        assert_eq!(project.chapters.len(), 3);

        // Verify numeric sorting (1, 2, 10) not alphabetical (1, 10, 2)
        assert_eq!(project.chapters[0].number, 1);
        assert_eq!(project.chapters[1].number, 2);
        assert_eq!(project.chapters[2].number, 10);

        fs::remove_dir_all(test_dir).unwrap();
    }
}
