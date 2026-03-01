## ADDED Requirements

### Requirement: LaTeX display math rendered as image

The system SHALL render display math blocks (`$$...$$`) as images (SVG preferred, PNG fallback) by invoking an external LaTeX toolchain, and embed the resulting image in the DOCX as a centered `DocElement::Image`.

#### Scenario: Display math with LaTeX available

- **WHEN** a display math block is encountered and `math.renderer` is `"image"` and the LaTeX toolchain is available
- **THEN** the system SHALL render the LaTeX to an SVG image, embed it as a centered image paragraph in the DOCX, and auto-size it from the SVG viewBox constrained to page width

#### Scenario: Display math without LaTeX available

- **WHEN** a display math block is encountered and `math.renderer` is `"image"` but the LaTeX toolchain is not found
- **THEN** the system SHALL fall back to OMML rendering and emit a warning to stderr

#### Scenario: Display math with OMML mode

- **WHEN** a display math block is encountered and `math.renderer` is `"omml"`
- **THEN** the system SHALL use the existing OMML rendering path (current behavior)

### Requirement: LaTeX inline math rendered as image

The system SHALL render inline math (`$...$`) as inline images by invoking the LaTeX toolchain, embedding the image inline within the text run.

#### Scenario: Inline math with LaTeX available

- **WHEN** inline math is encountered and `math.renderer` is `"image"` and the LaTeX toolchain is available
- **THEN** the system SHALL render the LaTeX to an SVG image and embed it as an inline image within the current paragraph run

#### Scenario: Inline math fallback to OMML

- **WHEN** inline math is encountered and `math.renderer` is `"image"` but the LaTeX toolchain is not found
- **THEN** the system SHALL fall back to inline OMML rendering and emit a warning to stderr

### Requirement: Math rendering configuration

The system SHALL support a `[math]` section in `md2docx.toml`:

- `renderer`: `"image"` or `"omml"` (default: `"image"`)

#### Scenario: Default renderer is image

- **WHEN** no `[math]` section exists in `md2docx.toml`
- **THEN** the system SHALL default to `renderer = "image"`

#### Scenario: Explicit OMML renderer

- **WHEN** `md2docx.toml` contains `[math]` with `renderer = "omml"`
- **THEN** the system SHALL use OMML for all math expressions

### Requirement: LaTeX toolchain detection

The system SHALL detect the availability of the LaTeX toolchain by checking for the `latex` and `dvisvgm` executables in the system PATH at startup.

#### Scenario: Toolchain present

- **WHEN** both `latex` and `dvisvgm` are found in PATH
- **THEN** the system SHALL mark the image renderer as available

#### Scenario: Toolchain missing

- **WHEN** `latex` or `dvisvgm` is not found in PATH
- **THEN** the system SHALL mark the image renderer as unavailable and log a warning suggesting installation

### Requirement: Equation image caching

The system SHALL cache rendered equation images within a single document build. If the same LaTeX expression appears multiple times, the cached image SHALL be reused.

#### Scenario: Duplicate expression

- **WHEN** the same LaTeX expression appears in two different locations in the document
- **THEN** the system SHALL render the expression once and reference the same image relationship ID for both occurrences

### Requirement: LaTeX rendering pipeline

The system SHALL render LaTeX expressions to SVG using the following pipeline:

1. Write expression to a temporary `.tex` file with a minimal document preamble
2. Run `latex` to produce a DVI file
3. Run `dvisvgm --exact` to produce a tightly-cropped SVG
4. Read the SVG and embed as image data

#### Scenario: Successful render pipeline

- **WHEN** a LaTeX expression is rendered
- **THEN** the system SHALL produce an SVG image with tight cropping (no excess whitespace) suitable for embedding in DOCX
