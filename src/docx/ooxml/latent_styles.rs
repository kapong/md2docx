//! Latent Style Catalog for OOXML
//!
//! This module provides Microsoft Word's complete latent style catalog (376 styles).
//! By including this catalog in generated documents, Word will recognize all built-in
//! styles without needing to "repair" the document or add missing style definitions.
//!
//! Ported from docxgo (github.com/mmonterroca/docxgo) internal/serializer/latent_styles.go

use quick_xml::events::{BytesStart, Event};
use quick_xml::Writer;
use std::io::Cursor;

use crate::error::Result;

/// A latent style exception defines metadata for a built-in Word style
#[derive(Debug, Clone, Copy)]
pub struct LatentStyleException {
    /// The style name (must match Word's built-in name exactly)
    pub name: &'static str,
    /// UI priority (controls order in style gallery). None means use default (99)
    pub ui_priority: Option<u32>,
    /// If true, style is hidden from UI until used
    pub semi_hidden: bool,
    /// If true, style becomes visible when used
    pub unhide_when_used: bool,
    /// If true, style appears in Quick Styles gallery
    pub q_format: bool,
}

impl LatentStyleException {
    /// Create a new latent style exception with all defaults (hidden, no qFormat)
    const fn new(name: &'static str) -> Self {
        Self {
            name,
            ui_priority: None,
            semi_hidden: false,
            unhide_when_used: false,
            q_format: false,
        }
    }

    /// Set UI priority
    const fn priority(mut self, p: u32) -> Self {
        self.ui_priority = Some(p);
        self
    }

    /// Mark as semi-hidden and unhide when used
    const fn hidden(mut self) -> Self {
        self.semi_hidden = true;
        self.unhide_when_used = true;
        self
    }

    /// Mark as semi-hidden only (no unhide when used)
    const fn semi_hidden_only(mut self) -> Self {
        self.semi_hidden = true;
        self
    }

    /// Mark as quick format (appears in Quick Styles gallery)
    const fn qformat(mut self) -> Self {
        self.q_format = true;
        self
    }
}

/// The latent styles container with defaults and exceptions
#[derive(Debug, Clone)]
pub struct LatentStyles {
    /// Default locked state for all styles
    pub def_locked_state: bool,
    /// Default UI priority (99 = low priority)
    pub def_ui_priority: u32,
    /// Default semi-hidden state
    pub def_semi_hidden: bool,
    /// Default unhide-when-used state
    pub def_unhide_when_used: bool,
    /// Default quick format state
    pub def_q_format: bool,
    /// Total count of latent styles
    pub count: u32,
    /// Style exceptions that override defaults
    pub exceptions: &'static [LatentStyleException],
}

impl Default for LatentStyles {
    fn default() -> Self {
        Self {
            def_locked_state: false,
            def_ui_priority: 99,
            def_semi_hidden: false,
            def_unhide_when_used: false,
            def_q_format: false,
            count: 376,
            exceptions: &DEFAULT_EXCEPTIONS,
        }
    }
}

impl LatentStyles {
    /// Generate XML for the latentStyles element
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        // Start latentStyles element with attributes
        let mut elem = BytesStart::new("w:latentStyles");
        elem.push_attribute((
            "w:defLockedState",
            if self.def_locked_state { "1" } else { "0" },
        ));
        elem.push_attribute(("w:defUIPriority", self.def_ui_priority.to_string().as_str()));
        elem.push_attribute((
            "w:defSemiHidden",
            if self.def_semi_hidden { "1" } else { "0" },
        ));
        elem.push_attribute((
            "w:defUnhideWhenUsed",
            if self.def_unhide_when_used { "1" } else { "0" },
        ));
        elem.push_attribute(("w:defQFormat", if self.def_q_format { "1" } else { "0" }));
        elem.push_attribute(("w:count", self.count.to_string().as_str()));
        writer.write_event(Event::Start(elem))?;

        // Write each exception
        for exc in self.exceptions {
            let mut exc_elem = BytesStart::new("w:lsdException");
            exc_elem.push_attribute(("w:name", exc.name));

            // Only include non-default attributes
            if let Some(priority) = exc.ui_priority {
                exc_elem.push_attribute(("w:uiPriority", priority.to_string().as_str()));
            }
            if exc.semi_hidden {
                exc_elem.push_attribute(("w:semiHidden", "1"));
            }
            if exc.unhide_when_used {
                exc_elem.push_attribute(("w:unhideWhenUsed", "1"));
            }
            if exc.q_format {
                exc_elem.push_attribute(("w:qFormat", "1"));
            }

            writer.write_event(Event::Empty(exc_elem))?;
        }

        // Close latentStyles
        writer.write_event(Event::End(quick_xml::events::BytesEnd::new(
            "w:latentStyles",
        )))?;

        Ok(writer.into_inner().into_inner())
    }
}

/// All 376 latent style exceptions matching Microsoft Word's built-in catalog
pub static DEFAULT_EXCEPTIONS: [LatentStyleException; 376] = [
    // Core styles
    LatentStyleException::new("Normal").priority(0).qformat(),
    LatentStyleException::new("heading 1").priority(9).qformat(),
    LatentStyleException::new("heading 2")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 3")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 4")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 5")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 6")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 7")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 8")
        .priority(9)
        .qformat()
        .hidden(),
    LatentStyleException::new("heading 9")
        .priority(9)
        .qformat()
        .hidden(),
    // Index styles
    LatentStyleException::new("index 1").hidden(),
    LatentStyleException::new("index 2").hidden(),
    LatentStyleException::new("index 3").hidden(),
    LatentStyleException::new("index 4").hidden(),
    LatentStyleException::new("index 5").hidden(),
    LatentStyleException::new("index 6").hidden(),
    LatentStyleException::new("index 7").hidden(),
    LatentStyleException::new("index 8").hidden(),
    LatentStyleException::new("index 9").hidden(),
    // TOC styles
    LatentStyleException::new("toc 1").priority(39).hidden(),
    LatentStyleException::new("toc 2").priority(39).hidden(),
    LatentStyleException::new("toc 3").priority(39).hidden(),
    LatentStyleException::new("toc 4").priority(39).hidden(),
    LatentStyleException::new("toc 5").priority(39).hidden(),
    LatentStyleException::new("toc 6").priority(39).hidden(),
    LatentStyleException::new("toc 7").priority(39).hidden(),
    LatentStyleException::new("toc 8").priority(39).hidden(),
    LatentStyleException::new("toc 9").priority(39).hidden(),
    // Document structure styles
    LatentStyleException::new("Normal Indent").hidden(),
    LatentStyleException::new("footnote text").hidden(),
    LatentStyleException::new("annotation text").hidden(),
    LatentStyleException::new("header").hidden(),
    LatentStyleException::new("footer").hidden(),
    LatentStyleException::new("index heading").hidden(),
    LatentStyleException::new("caption")
        .priority(35)
        .qformat()
        .hidden(),
    LatentStyleException::new("table of figures").hidden(),
    LatentStyleException::new("envelope address").hidden(),
    LatentStyleException::new("envelope return").hidden(),
    LatentStyleException::new("footnote reference").hidden(),
    LatentStyleException::new("annotation reference").hidden(),
    LatentStyleException::new("line number").hidden(),
    LatentStyleException::new("page number").hidden(),
    LatentStyleException::new("endnote reference").hidden(),
    LatentStyleException::new("endnote text").hidden(),
    LatentStyleException::new("table of authorities").hidden(),
    LatentStyleException::new("macro").hidden(),
    LatentStyleException::new("toa heading").hidden(),
    LatentStyleException::new("List").hidden(),
    LatentStyleException::new("List Bullet").hidden(),
    LatentStyleException::new("List Number").hidden(),
    LatentStyleException::new("List 2").hidden(),
    LatentStyleException::new("List 3").hidden(),
    LatentStyleException::new("List 4").hidden(),
    LatentStyleException::new("List 5").hidden(),
    LatentStyleException::new("List Bullet 2").hidden(),
    LatentStyleException::new("List Bullet 3").hidden(),
    LatentStyleException::new("List Bullet 4").hidden(),
    LatentStyleException::new("List Bullet 5").hidden(),
    LatentStyleException::new("List Number 2").hidden(),
    LatentStyleException::new("List Number 3").hidden(),
    LatentStyleException::new("List Number 4").hidden(),
    LatentStyleException::new("List Number 5").hidden(),
    LatentStyleException::new("Title").priority(10).qformat(),
    LatentStyleException::new("Closing").hidden(),
    LatentStyleException::new("Signature").hidden(),
    LatentStyleException::new("Default Paragraph Font")
        .priority(1)
        .hidden(),
    LatentStyleException::new("Body Text").hidden(),
    LatentStyleException::new("Body Text Indent").hidden(),
    LatentStyleException::new("List Continue").hidden(),
    LatentStyleException::new("List Continue 2").hidden(),
    LatentStyleException::new("List Continue 3").hidden(),
    LatentStyleException::new("List Continue 4").hidden(),
    LatentStyleException::new("List Continue 5").hidden(),
    LatentStyleException::new("Message Header").hidden(),
    LatentStyleException::new("Subtitle").priority(11).qformat(),
    LatentStyleException::new("Salutation").hidden(),
    LatentStyleException::new("Date").hidden(),
    LatentStyleException::new("Body Text First Indent").hidden(),
    LatentStyleException::new("Body Text First Indent 2").hidden(),
    LatentStyleException::new("Note Heading").hidden(),
    LatentStyleException::new("Body Text 2").hidden(),
    LatentStyleException::new("Body Text 3").hidden(),
    LatentStyleException::new("Body Text Indent 2").hidden(),
    LatentStyleException::new("Body Text Indent 3").hidden(),
    LatentStyleException::new("Block Text").hidden(),
    LatentStyleException::new("Hyperlink").hidden(),
    LatentStyleException::new("FollowedHyperlink").hidden(),
    LatentStyleException::new("Strong").priority(22).qformat(),
    LatentStyleException::new("Emphasis").priority(20).qformat(),
    LatentStyleException::new("Document Map").hidden(),
    LatentStyleException::new("Plain Text").hidden(),
    LatentStyleException::new("E-mail Signature").hidden(),
    LatentStyleException::new("HTML Top of Form").hidden(),
    LatentStyleException::new("HTML Bottom of Form").hidden(),
    LatentStyleException::new("Normal (Web)").hidden(),
    LatentStyleException::new("HTML Acronym").hidden(),
    LatentStyleException::new("HTML Address").hidden(),
    LatentStyleException::new("HTML Cite").hidden(),
    LatentStyleException::new("HTML Code").hidden(),
    LatentStyleException::new("HTML Definition").hidden(),
    LatentStyleException::new("HTML Keyboard").hidden(),
    LatentStyleException::new("HTML Preformatted").hidden(),
    LatentStyleException::new("HTML Sample").hidden(),
    LatentStyleException::new("HTML Typewriter").hidden(),
    LatentStyleException::new("HTML Variable").hidden(),
    LatentStyleException::new("Normal Table").hidden(),
    LatentStyleException::new("annotation subject").hidden(),
    LatentStyleException::new("No List").hidden(),
    LatentStyleException::new("Outline List 1").hidden(),
    LatentStyleException::new("Outline List 2").hidden(),
    LatentStyleException::new("Outline List 3").hidden(),
    LatentStyleException::new("Table Simple 1").hidden(),
    LatentStyleException::new("Table Simple 2").hidden(),
    LatentStyleException::new("Table Simple 3").hidden(),
    LatentStyleException::new("Table Classic 1").hidden(),
    LatentStyleException::new("Table Classic 2").hidden(),
    LatentStyleException::new("Table Classic 3").hidden(),
    LatentStyleException::new("Table Classic 4").hidden(),
    LatentStyleException::new("Table Colorful 1").hidden(),
    LatentStyleException::new("Table Colorful 2").hidden(),
    LatentStyleException::new("Table Colorful 3").hidden(),
    LatentStyleException::new("Table Columns 1").hidden(),
    LatentStyleException::new("Table Columns 2").hidden(),
    LatentStyleException::new("Table Columns 3").hidden(),
    LatentStyleException::new("Table Columns 4").hidden(),
    LatentStyleException::new("Table Columns 5").hidden(),
    LatentStyleException::new("Table Grid 1").hidden(),
    LatentStyleException::new("Table Grid 2").hidden(),
    LatentStyleException::new("Table Grid 3").hidden(),
    LatentStyleException::new("Table Grid 4").hidden(),
    LatentStyleException::new("Table Grid 5").hidden(),
    LatentStyleException::new("Table Grid 6").hidden(),
    LatentStyleException::new("Table Grid 7").hidden(),
    LatentStyleException::new("Table Grid 8").hidden(),
    LatentStyleException::new("Table List 1").hidden(),
    LatentStyleException::new("Table List 2").hidden(),
    LatentStyleException::new("Table List 3").hidden(),
    LatentStyleException::new("Table List 4").hidden(),
    LatentStyleException::new("Table List 5").hidden(),
    LatentStyleException::new("Table List 6").hidden(),
    LatentStyleException::new("Table List 7").hidden(),
    LatentStyleException::new("Table List 8").hidden(),
    LatentStyleException::new("Table 3D effects 1").hidden(),
    LatentStyleException::new("Table 3D effects 2").hidden(),
    LatentStyleException::new("Table 3D effects 3").hidden(),
    LatentStyleException::new("Table Contemporary").hidden(),
    LatentStyleException::new("Table Elegant").hidden(),
    LatentStyleException::new("Table Professional").hidden(),
    LatentStyleException::new("Table Subtle 1").hidden(),
    LatentStyleException::new("Table Subtle 2").hidden(),
    LatentStyleException::new("Table Web 1").hidden(),
    LatentStyleException::new("Table Web 2").hidden(),
    LatentStyleException::new("Table Web 3").hidden(),
    LatentStyleException::new("Balloon Text").hidden(),
    LatentStyleException::new("Table Grid").priority(39),
    LatentStyleException::new("Table Theme").hidden(),
    LatentStyleException::new("Placeholder Text").semi_hidden_only(),
    LatentStyleException::new("No Spacing")
        .priority(1)
        .qformat(),
    // Table styles - base colors
    LatentStyleException::new("Light Shading").priority(60),
    LatentStyleException::new("Light List").priority(61),
    LatentStyleException::new("Light Grid").priority(62),
    LatentStyleException::new("Medium Shading 1").priority(63),
    LatentStyleException::new("Medium Shading 2").priority(64),
    LatentStyleException::new("Medium List 1").priority(65),
    LatentStyleException::new("Medium List 2").priority(66),
    LatentStyleException::new("Medium Grid 1").priority(67),
    LatentStyleException::new("Medium Grid 2").priority(68),
    LatentStyleException::new("Medium Grid 3").priority(69),
    LatentStyleException::new("Dark List").priority(70),
    LatentStyleException::new("Colorful Shading").priority(71),
    LatentStyleException::new("Colorful List").priority(72),
    LatentStyleException::new("Colorful Grid").priority(73),
    // Table styles - Accent 1
    LatentStyleException::new("Light Shading Accent 1").priority(60),
    LatentStyleException::new("Light List Accent 1").priority(61),
    LatentStyleException::new("Light Grid Accent 1").priority(62),
    LatentStyleException::new("Medium Shading 1 Accent 1").priority(63),
    LatentStyleException::new("Medium Shading 2 Accent 1").priority(64),
    LatentStyleException::new("Medium List 1 Accent 1").priority(65),
    LatentStyleException::new("Revision").semi_hidden_only(),
    LatentStyleException::new("List Paragraph")
        .priority(34)
        .qformat(),
    LatentStyleException::new("Quote").priority(29).qformat(),
    LatentStyleException::new("Intense Quote")
        .priority(30)
        .qformat(),
    LatentStyleException::new("Medium List 2 Accent 1").priority(66),
    LatentStyleException::new("Medium Grid 1 Accent 1").priority(67),
    LatentStyleException::new("Medium Grid 2 Accent 1").priority(68),
    LatentStyleException::new("Medium Grid 3 Accent 1").priority(69),
    LatentStyleException::new("Dark List Accent 1").priority(70),
    LatentStyleException::new("Colorful Shading Accent 1").priority(71),
    LatentStyleException::new("Colorful List Accent 1").priority(72),
    LatentStyleException::new("Colorful Grid Accent 1").priority(73),
    // Table styles - Accent 2
    LatentStyleException::new("Light Shading Accent 2").priority(60),
    LatentStyleException::new("Light List Accent 2").priority(61),
    LatentStyleException::new("Light Grid Accent 2").priority(62),
    LatentStyleException::new("Medium Shading 1 Accent 2").priority(63),
    LatentStyleException::new("Medium Shading 2 Accent 2").priority(64),
    LatentStyleException::new("Medium List 1 Accent 2").priority(65),
    LatentStyleException::new("Medium List 2 Accent 2").priority(66),
    LatentStyleException::new("Medium Grid 1 Accent 2").priority(67),
    LatentStyleException::new("Medium Grid 2 Accent 2").priority(68),
    LatentStyleException::new("Medium Grid 3 Accent 2").priority(69),
    LatentStyleException::new("Dark List Accent 2").priority(70),
    LatentStyleException::new("Colorful Shading Accent 2").priority(71),
    LatentStyleException::new("Colorful List Accent 2").priority(72),
    LatentStyleException::new("Colorful Grid Accent 2").priority(73),
    // Table styles - Accent 3
    LatentStyleException::new("Light Shading Accent 3").priority(60),
    LatentStyleException::new("Light List Accent 3").priority(61),
    LatentStyleException::new("Light Grid Accent 3").priority(62),
    LatentStyleException::new("Medium Shading 1 Accent 3").priority(63),
    LatentStyleException::new("Medium Shading 2 Accent 3").priority(64),
    LatentStyleException::new("Medium List 1 Accent 3").priority(65),
    LatentStyleException::new("Medium List 2 Accent 3").priority(66),
    LatentStyleException::new("Medium Grid 1 Accent 3").priority(67),
    LatentStyleException::new("Medium Grid 2 Accent 3").priority(68),
    LatentStyleException::new("Medium Grid 3 Accent 3").priority(69),
    LatentStyleException::new("Dark List Accent 3").priority(70),
    LatentStyleException::new("Colorful Shading Accent 3").priority(71),
    LatentStyleException::new("Colorful List Accent 3").priority(72),
    LatentStyleException::new("Colorful Grid Accent 3").priority(73),
    // Table styles - Accent 4
    LatentStyleException::new("Light Shading Accent 4").priority(60),
    LatentStyleException::new("Light List Accent 4").priority(61),
    LatentStyleException::new("Light Grid Accent 4").priority(62),
    LatentStyleException::new("Medium Shading 1 Accent 4").priority(63),
    LatentStyleException::new("Medium Shading 2 Accent 4").priority(64),
    LatentStyleException::new("Medium List 1 Accent 4").priority(65),
    LatentStyleException::new("Medium List 2 Accent 4").priority(66),
    LatentStyleException::new("Medium Grid 1 Accent 4").priority(67),
    LatentStyleException::new("Medium Grid 2 Accent 4").priority(68),
    LatentStyleException::new("Medium Grid 3 Accent 4").priority(69),
    LatentStyleException::new("Dark List Accent 4").priority(70),
    LatentStyleException::new("Colorful Shading Accent 4").priority(71),
    LatentStyleException::new("Colorful List Accent 4").priority(72),
    LatentStyleException::new("Colorful Grid Accent 4").priority(73),
    // Table styles - Accent 5
    LatentStyleException::new("Light Shading Accent 5").priority(60),
    LatentStyleException::new("Light List Accent 5").priority(61),
    LatentStyleException::new("Light Grid Accent 5").priority(62),
    LatentStyleException::new("Medium Shading 1 Accent 5").priority(63),
    LatentStyleException::new("Medium Shading 2 Accent 5").priority(64),
    LatentStyleException::new("Medium List 1 Accent 5").priority(65),
    LatentStyleException::new("Medium List 2 Accent 5").priority(66),
    LatentStyleException::new("Medium Grid 1 Accent 5").priority(67),
    LatentStyleException::new("Medium Grid 2 Accent 5").priority(68),
    LatentStyleException::new("Medium Grid 3 Accent 5").priority(69),
    LatentStyleException::new("Dark List Accent 5").priority(70),
    LatentStyleException::new("Colorful Shading Accent 5").priority(71),
    LatentStyleException::new("Colorful List Accent 5").priority(72),
    LatentStyleException::new("Colorful Grid Accent 5").priority(73),
    // Table styles - Accent 6
    LatentStyleException::new("Light Shading Accent 6").priority(60),
    LatentStyleException::new("Light List Accent 6").priority(61),
    LatentStyleException::new("Light Grid Accent 6").priority(62),
    LatentStyleException::new("Medium Shading 1 Accent 6").priority(63),
    LatentStyleException::new("Medium Shading 2 Accent 6").priority(64),
    LatentStyleException::new("Medium List 1 Accent 6").priority(65),
    LatentStyleException::new("Medium List 2 Accent 6").priority(66),
    LatentStyleException::new("Medium Grid 1 Accent 6").priority(67),
    LatentStyleException::new("Medium Grid 2 Accent 6").priority(68),
    LatentStyleException::new("Medium Grid 3 Accent 6").priority(69),
    LatentStyleException::new("Dark List Accent 6").priority(70),
    LatentStyleException::new("Colorful Shading Accent 6").priority(71),
    LatentStyleException::new("Colorful List Accent 6").priority(72),
    LatentStyleException::new("Colorful Grid Accent 6").priority(73),
    // Emphasis styles
    LatentStyleException::new("Subtle Emphasis")
        .priority(19)
        .qformat(),
    LatentStyleException::new("Intense Emphasis")
        .priority(21)
        .qformat(),
    LatentStyleException::new("Subtle Reference")
        .priority(31)
        .qformat(),
    LatentStyleException::new("Intense Reference")
        .priority(32)
        .qformat(),
    LatentStyleException::new("Book Title")
        .priority(33)
        .qformat(),
    LatentStyleException::new("Bibliography")
        .priority(37)
        .hidden(),
    LatentStyleException::new("TOC Heading")
        .priority(39)
        .qformat()
        .hidden(),
    // Plain Table styles
    LatentStyleException::new("Plain Table 1").priority(41),
    LatentStyleException::new("Plain Table 2").priority(42),
    LatentStyleException::new("Plain Table 3").priority(43),
    LatentStyleException::new("Plain Table 4").priority(44),
    LatentStyleException::new("Plain Table 5").priority(45),
    // Grid Table styles
    LatentStyleException::new("Grid Table Light").priority(40),
    LatentStyleException::new("Grid Table 1 Light").priority(46),
    LatentStyleException::new("Grid Table 2").priority(47),
    LatentStyleException::new("Grid Table 3").priority(48),
    LatentStyleException::new("Grid Table 4").priority(49),
    LatentStyleException::new("Grid Table 5 Dark").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful").priority(52),
    // Grid Table - Accent 1
    LatentStyleException::new("Grid Table 1 Light Accent 1").priority(46),
    LatentStyleException::new("Grid Table 2 Accent 1").priority(47),
    LatentStyleException::new("Grid Table 3 Accent 1").priority(48),
    LatentStyleException::new("Grid Table 4 Accent 1").priority(49),
    LatentStyleException::new("Grid Table 5 Dark Accent 1").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful Accent 1").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful Accent 1").priority(52),
    // Grid Table - Accent 2
    LatentStyleException::new("Grid Table 1 Light Accent 2").priority(46),
    LatentStyleException::new("Grid Table 2 Accent 2").priority(47),
    LatentStyleException::new("Grid Table 3 Accent 2").priority(48),
    LatentStyleException::new("Grid Table 4 Accent 2").priority(49),
    LatentStyleException::new("Grid Table 5 Dark Accent 2").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful Accent 2").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful Accent 2").priority(52),
    // Grid Table - Accent 3
    LatentStyleException::new("Grid Table 1 Light Accent 3").priority(46),
    LatentStyleException::new("Grid Table 2 Accent 3").priority(47),
    LatentStyleException::new("Grid Table 3 Accent 3").priority(48),
    LatentStyleException::new("Grid Table 4 Accent 3").priority(49),
    LatentStyleException::new("Grid Table 5 Dark Accent 3").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful Accent 3").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful Accent 3").priority(52),
    // Grid Table - Accent 4
    LatentStyleException::new("Grid Table 1 Light Accent 4").priority(46),
    LatentStyleException::new("Grid Table 2 Accent 4").priority(47),
    LatentStyleException::new("Grid Table 3 Accent 4").priority(48),
    LatentStyleException::new("Grid Table 4 Accent 4").priority(49),
    LatentStyleException::new("Grid Table 5 Dark Accent 4").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful Accent 4").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful Accent 4").priority(52),
    // Grid Table - Accent 5
    LatentStyleException::new("Grid Table 1 Light Accent 5").priority(46),
    LatentStyleException::new("Grid Table 2 Accent 5").priority(47),
    LatentStyleException::new("Grid Table 3 Accent 5").priority(48),
    LatentStyleException::new("Grid Table 4 Accent 5").priority(49),
    LatentStyleException::new("Grid Table 5 Dark Accent 5").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful Accent 5").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful Accent 5").priority(52),
    // Grid Table - Accent 6
    LatentStyleException::new("Grid Table 1 Light Accent 6").priority(46),
    LatentStyleException::new("Grid Table 2 Accent 6").priority(47),
    LatentStyleException::new("Grid Table 3 Accent 6").priority(48),
    LatentStyleException::new("Grid Table 4 Accent 6").priority(49),
    LatentStyleException::new("Grid Table 5 Dark Accent 6").priority(50),
    LatentStyleException::new("Grid Table 6 Colorful Accent 6").priority(51),
    LatentStyleException::new("Grid Table 7 Colorful Accent 6").priority(52),
    // List Table styles - base
    LatentStyleException::new("List Table 1 Light").priority(46),
    LatentStyleException::new("List Table 2").priority(47),
    LatentStyleException::new("List Table 3").priority(48),
    LatentStyleException::new("List Table 4").priority(49),
    LatentStyleException::new("List Table 5 Dark").priority(50),
    LatentStyleException::new("List Table 6 Colorful").priority(51),
    LatentStyleException::new("List Table 7 Colorful").priority(52),
    // List Table - Accent 1
    LatentStyleException::new("List Table 1 Light Accent 1").priority(46),
    LatentStyleException::new("List Table 2 Accent 1").priority(47),
    LatentStyleException::new("List Table 3 Accent 1").priority(48),
    LatentStyleException::new("List Table 4 Accent 1").priority(49),
    LatentStyleException::new("List Table 5 Dark Accent 1").priority(50),
    LatentStyleException::new("List Table 6 Colorful Accent 1").priority(51),
    LatentStyleException::new("List Table 7 Colorful Accent 1").priority(52),
    // List Table - Accent 2
    LatentStyleException::new("List Table 1 Light Accent 2").priority(46),
    LatentStyleException::new("List Table 2 Accent 2").priority(47),
    LatentStyleException::new("List Table 3 Accent 2").priority(48),
    LatentStyleException::new("List Table 4 Accent 2").priority(49),
    LatentStyleException::new("List Table 5 Dark Accent 2").priority(50),
    LatentStyleException::new("List Table 6 Colorful Accent 2").priority(51),
    LatentStyleException::new("List Table 7 Colorful Accent 2").priority(52),
    // List Table - Accent 3
    LatentStyleException::new("List Table 1 Light Accent 3").priority(46),
    LatentStyleException::new("List Table 2 Accent 3").priority(47),
    LatentStyleException::new("List Table 3 Accent 3").priority(48),
    LatentStyleException::new("List Table 4 Accent 3").priority(49),
    LatentStyleException::new("List Table 5 Dark Accent 3").priority(50),
    LatentStyleException::new("List Table 6 Colorful Accent 3").priority(51),
    LatentStyleException::new("List Table 7 Colorful Accent 3").priority(52),
    // List Table - Accent 4
    LatentStyleException::new("List Table 1 Light Accent 4").priority(46),
    LatentStyleException::new("List Table 2 Accent 4").priority(47),
    LatentStyleException::new("List Table 3 Accent 4").priority(48),
    LatentStyleException::new("List Table 4 Accent 4").priority(49),
    LatentStyleException::new("List Table 5 Dark Accent 4").priority(50),
    LatentStyleException::new("List Table 6 Colorful Accent 4").priority(51),
    LatentStyleException::new("List Table 7 Colorful Accent 4").priority(52),
    // List Table - Accent 5
    LatentStyleException::new("List Table 1 Light Accent 5").priority(46),
    LatentStyleException::new("List Table 2 Accent 5").priority(47),
    LatentStyleException::new("List Table 3 Accent 5").priority(48),
    LatentStyleException::new("List Table 4 Accent 5").priority(49),
    LatentStyleException::new("List Table 5 Dark Accent 5").priority(50),
    LatentStyleException::new("List Table 6 Colorful Accent 5").priority(51),
    LatentStyleException::new("List Table 7 Colorful Accent 5").priority(52),
    // List Table - Accent 6
    LatentStyleException::new("List Table 1 Light Accent 6").priority(46),
    LatentStyleException::new("List Table 2 Accent 6").priority(47),
    LatentStyleException::new("List Table 3 Accent 6").priority(48),
    LatentStyleException::new("List Table 4 Accent 6").priority(49),
    LatentStyleException::new("List Table 5 Dark Accent 6").priority(50),
    LatentStyleException::new("List Table 6 Colorful Accent 6").priority(51),
    LatentStyleException::new("List Table 7 Colorful Accent 6").priority(52),
    // Modern styles (Office 2016+)
    LatentStyleException::new("Mention").hidden(),
    LatentStyleException::new("Smart Hyperlink").hidden(),
    LatentStyleException::new("Hashtag").hidden(),
    LatentStyleException::new("Unresolved Mention").hidden(),
    LatentStyleException::new("Smart Link").hidden(),
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_exceptions_count() {
        assert_eq!(DEFAULT_EXCEPTIONS.len(), 376);
    }

    #[test]
    fn test_latent_styles_default() {
        let styles = LatentStyles::default();
        assert_eq!(styles.count, 376);
        assert_eq!(styles.def_ui_priority, 99);
        assert!(!styles.def_locked_state);
        assert!(!styles.def_semi_hidden);
        assert!(!styles.def_unhide_when_used);
        assert!(!styles.def_q_format);
    }

    #[test]
    fn test_latent_styles_to_xml() {
        let styles = LatentStyles::default();
        let xml = styles.to_xml().expect("Failed to generate XML");
        let xml_str = String::from_utf8(xml).expect("Invalid UTF-8");

        // Check root element
        assert!(xml_str.contains("<w:latentStyles"));
        assert!(xml_str.contains("w:defLockedState=\"0\""));
        assert!(xml_str.contains("w:defUIPriority=\"99\""));
        assert!(xml_str.contains("w:count=\"376\""));

        // Check some specific exceptions
        assert!(xml_str.contains("w:name=\"Normal\""));
        assert!(xml_str.contains("w:name=\"heading 1\""));
        assert!(xml_str.contains("w:name=\"Title\""));
    }

    #[test]
    fn test_first_styles_correct() {
        // Verify the first few styles match Word's catalog
        assert_eq!(DEFAULT_EXCEPTIONS[0].name, "Normal");
        assert_eq!(DEFAULT_EXCEPTIONS[0].ui_priority, Some(0));
        assert!(DEFAULT_EXCEPTIONS[0].q_format);

        assert_eq!(DEFAULT_EXCEPTIONS[1].name, "heading 1");
        assert_eq!(DEFAULT_EXCEPTIONS[1].ui_priority, Some(9));
        assert!(DEFAULT_EXCEPTIONS[1].q_format);

        assert_eq!(DEFAULT_EXCEPTIONS[2].name, "heading 2");
        assert!(DEFAULT_EXCEPTIONS[2].semi_hidden);
        assert!(DEFAULT_EXCEPTIONS[2].unhide_when_used);
    }
}
