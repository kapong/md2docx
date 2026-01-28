//! Script detection for Thai/English text
//!
//! This module provides functions to detect the script of text,
//! primarily for distinguishing Thai from Latin/English text.

/// Thai Unicode range: U+0E00 to U+0E7F
const THAI_START: char = '\u{0E00}';
const THAI_END: char = '\u{0E7F}';

/// Check if a character is Thai
#[inline]
pub fn is_thai_char(c: char) -> bool {
    c >= THAI_START && c <= THAI_END
}

/// Check if a string contains any Thai characters
pub fn contains_thai(text: &str) -> bool {
    text.chars().any(is_thai_char)
}

/// Check if a string is predominantly Thai (>50% Thai characters, excluding spaces/punctuation)
pub fn is_predominantly_thai(text: &str) -> bool {
    let mut thai_count = 0;
    let mut letter_count = 0;

    for c in text.chars() {
        if c.is_alphabetic() {
            letter_count += 1;
            if is_thai_char(c) {
                thai_count += 1;
            }
        }
    }

    if letter_count == 0 {
        return false;
    }

    thai_count * 2 > letter_count // More than 50% Thai
}

/// Determine the primary language of a text for spell-check purposes
/// Returns "th-TH" for Thai, "en-US" for English/Latin
pub fn detect_language(text: &str) -> &'static str {
    if is_predominantly_thai(text) {
        "th-TH"
    } else {
        "en-US"
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_is_thai_char() {
        assert!(is_thai_char('ก')); // Thai Ko Kai
        assert!(is_thai_char('ข')); // Thai Kho Khai
        assert!(is_thai_char('๐')); // Thai digit zero
        assert!(is_thai_char('๙')); // Thai digit nine
        assert!(!is_thai_char('a'));
        assert!(!is_thai_char('Z'));
        assert!(!is_thai_char('1'));
        assert!(!is_thai_char(' '));
    }

    #[test]
    fn test_contains_thai() {
        assert!(contains_thai("สวัสดี"));
        assert!(contains_thai("Hello สวัสดี World"));
        assert!(contains_thai("Mixed ภาษา text"));
        assert!(!contains_thai("Hello World"));
        assert!(!contains_thai("12345"));
        assert!(!contains_thai(""));
    }

    #[test]
    fn test_is_predominantly_thai() {
        assert!(is_predominantly_thai("สวัสดีครับ"));
        assert!(is_predominantly_thai("นี่คือข้อความภาษาไทย"));
        assert!(!is_predominantly_thai("Hello World"));
        // "Hello สวัสดี" = 5 English + 6 Thai = Thai-heavy (6 > 5)
        assert!(is_predominantly_thai("Hello สวัสดี"));
        assert!(is_predominantly_thai("สวัสดี Hello"));
        assert!(!is_predominantly_thai("")); // Empty
        assert!(!is_predominantly_thai("12345")); // Numbers only
                                                  // More English than Thai
        assert!(!is_predominantly_thai("Hello World สวัสดี"));
    }

    #[test]
    fn test_detect_language() {
        assert_eq!(detect_language("สวัสดีครับ"), "th-TH");
        assert_eq!(detect_language("Hello World"), "en-US");
        assert_eq!(detect_language("สวัสดี Hello"), "th-TH"); // Thai-heavy
        assert_eq!(detect_language("Hello สวัสดี World"), "en-US"); // English-heavy
    }
}
