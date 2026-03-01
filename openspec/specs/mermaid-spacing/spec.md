## ADDED Requirements

### Requirement: Mermaid diagram vertical spacing

The system SHALL add configurable vertical spacing before and after mermaid diagram image paragraphs in the DOCX output. The spacing SHALL be applied as OOXML paragraph spacing properties (`w:spacing w:before` / `w:after`).

#### Scenario: Default spacing applied

- **WHEN** a mermaid diagram is rendered and no `[mermaid]` config section exists
- **THEN** the diagram paragraph SHALL have 120 twips spacing before and 120 twips spacing after

#### Scenario: Custom spacing from configuration

- **WHEN** the `md2docx.toml` contains `[mermaid]` with `spacing_before = "240"` and `spacing_after = "60"`
- **THEN** the diagram paragraph SHALL use 240 twips before and 60 twips after

#### Scenario: Zero spacing configured

- **WHEN** the `md2docx.toml` contains `[mermaid]` with `spacing_before = "0"` and `spacing_after = "0"`
- **THEN** the diagram paragraph SHALL have no additional spacing (preserving legacy behavior)

### Requirement: Spacing applies only to image paragraph

The spacing SHALL be applied to the paragraph containing the mermaid image element. It SHALL NOT affect the caption paragraph spacing (which is controlled by the image template caption settings).

#### Scenario: Caption spacing unaffected

- **WHEN** a mermaid diagram has a caption and custom mermaid spacing is configured
- **THEN** the caption paragraph SHALL retain its own spacing settings from the image template
- **AND** the image paragraph SHALL use the mermaid spacing settings

### Requirement: Mermaid configuration section

The system SHALL support an optional `[mermaid]` section in `md2docx.toml` with the following fields:

- `spacing_before`: spacing before the diagram paragraph in twips (default: `"120"`)
- `spacing_after`: spacing after the diagram paragraph in twips (default: `"120"`)

#### Scenario: Configuration section parsed

- **WHEN** `md2docx.toml` contains a `[mermaid]` section
- **THEN** the system SHALL parse `spacing_before` and `spacing_after` as string values representing twips
