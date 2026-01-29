use std::collections::{HashMap, HashSet};

/// Manages Relationship IDs (rId) for the document.
///
/// This ensures all relationships in `word/_rels/document.xml.rels` have unique IDs,
/// even when combining content from multiple sources (main document, cover template, etc.).
#[derive(Debug, Clone)]
pub struct RelIdManager {
    next_id: usize,
    reserved_ids: HashSet<String>,
    // Mapping from (scope, original_id) -> new_id
    // Scope allows differentiating between cover.docx rIds and other sources
    mappings: HashMap<(String, String), String>,
}

impl RelIdManager {
    /// Create a new RelIdManager with standard IDs reserved
    pub fn new() -> Self {
        let mut manager = Self {
            next_id: 1,
            reserved_ids: HashSet::new(),
            mappings: HashMap::new(),
        };

        // Reserve standard IDs used in Relationships::document_rels()
        // rId1: styles.xml
        // rId2: settings.xml
        // rId3: fontTable.xml
        // rId4: webSettings.xml
        // rId5: theme/theme1.xml
        manager.reserve("rId1");
        manager.reserve("rId2");
        manager.reserve("rId3");
        manager.reserve("rId4");
        manager.reserve("rId5");

        manager
    }

    /// Reserve a specific ID (e.g. if a template uses a specific ID)
    pub fn reserve(&mut self, id: &str) {
        self.reserved_ids.insert(id.to_string());

        // Update next_id if reserved id is higher or equal
        if let Some(num_part) = id.strip_prefix("rId") {
            if let Ok(num) = num_part.parse::<usize>() {
                if num >= self.next_id {
                    self.next_id = num + 1;
                }
            }
        }
    }

    /// Generate a new unique rId
    pub fn next_id(&mut self) -> String {
        loop {
            let id = format!("rId{}", self.next_id);
            self.next_id += 1;
            if !self.reserved_ids.contains(&id) {
                // Mark as used (implicitly by incrementing next_id, but good to be safe if we change logic)
                return id;
            }
        }
    }

    /// Get a mapped ID for a template resource, creating one if needed.
    ///
    /// # Arguments
    /// * `scope` - A namespace for the source (e.g., "cover", "header1")
    /// * `original_id` - The rId in the source document (e.g., "rId7")
    pub fn get_mapped_id(&mut self, scope: &str, original_id: &str) -> String {
        let key = (scope.to_string(), original_id.to_string());
        if let Some(new_id) = self.mappings.get(&key) {
            return new_id.clone();
        }

        let new_id = self.next_id();
        self.mappings.insert(key, new_id.clone());
        new_id
    }

    /// Reset the manager (clearing mappings but keeping reserved IDs)
    pub fn reset(&mut self) {
        self.mappings.clear();
        // We don't reset next_id or reserved_ids to ensure safety if reused?
        // Actually, usually we want a fresh start for a new document.
        // But for multiple passes on same doc, we keep it.
    }
}

impl Default for RelIdManager {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_rel_id_manager() {
        let mut mgr = RelIdManager::new();

        // Should start after rId5
        let id1 = mgr.next_id();
        assert_eq!(id1, "rId6");

        let id2 = mgr.next_id();
        assert_eq!(id2, "rId7");
    }

    #[test]
    fn test_reserve() {
        let mut mgr = RelIdManager::new();

        // Test 1: Reserve an ID higher than current next_id bumps the counter
        mgr.reserve("rId10");
        // Reserving rId10 bumps next_id to 11, so next call returns rId11
        let id = mgr.next_id();
        assert_eq!(id, "rId11");

        // Test 2: Reserved IDs in sequence are skipped
        let mut mgr2 = RelIdManager::new();
        // After construction, next_id is 6 (rId1-5 reserved)
        let id1 = mgr2.next_id(); // rId6
        assert_eq!(id1, "rId6");

        mgr2.reserve("rId7"); // Reserve next expected ID
        let id2 = mgr2.next_id(); // Should skip rId7, return rId8
        assert_eq!(id2, "rId8");
    }

    #[test]
    fn test_mapping() {
        let mut mgr = RelIdManager::new();

        let new_id1 = mgr.get_mapped_id("cover", "rId1");
        let new_id2 = mgr.get_mapped_id("cover", "rId1"); // Same call

        assert_eq!(new_id1, new_id2);
        assert_ne!(new_id1, "rId1"); // Should be remapped to avoid conflict with reserved rId1

        let new_id3 = mgr.get_mapped_id("other", "rId1"); // Different scope
        assert_ne!(new_id1, new_id3);
    }
}
