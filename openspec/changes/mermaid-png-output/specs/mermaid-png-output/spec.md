## ADDED Requirements

### Requirement: Mermaid diagrams rendered as PNG by default

The system SHALL render Mermaid diagrams as PNG images at 150 DPI by default, instead of SVG, for universal Word compatibility.

#### Scenario: Default PNG output

- **WHEN** a mermaid diagram is rendered and no `output_format` is configured
- **THEN** the system SHALL produce a PNG image at 150 DPI and embed it in the DOCX

#### Scenario: Explicit PNG output

- **WHEN** the `md2docx.toml` contains `[mermaid]` with `output_format = "png"`
- **THEN** the system SHALL produce a PNG image and embed it in the DOCX

#### Scenario: SVG output opt-in

- **WHEN** the `md2docx.toml` contains `[mermaid]` with `output_format = "svg"`
- **THEN** the system SHALL produce an SVG image and embed it in the DOCX (legacy behavior)

### Requirement: Configurable DPI for PNG output

The system SHALL support a `dpi` field in the `[mermaid]` configuration section to control PNG image resolution. The DPI value SHALL be used to calculate the rendering scale factor and the physical dimensions of the embedded image.

#### Scenario: Default DPI

- **WHEN** a mermaid diagram is rendered as PNG and no `dpi` is configured
- **THEN** the system SHALL use 150 DPI

#### Scenario: Custom DPI

- **WHEN** the `md2docx.toml` contains `[mermaid]` with `dpi = 300`
- **THEN** the system SHALL render the PNG at 300 DPI

#### Scenario: DPI ignored for SVG

- **WHEN** the `md2docx.toml` contains `[mermaid]` with `output_format = "svg"` and `dpi = 300`
- **THEN** the system SHALL ignore the `dpi` setting and produce SVG output

### Requirement: PNG image dimensions in DOCX

The system SHALL calculate DOCX image dimensions (in EMU) from the PNG pixel dimensions and the configured DPI so that the image appears at the correct physical size in Word.

#### Scenario: Correct physical size at 150 DPI

- **WHEN** a mermaid PNG is 600px wide at 150 DPI
- **THEN** the embedded image width SHALL be 4 inches (600/150 * 914400 = 3657600 EMU)

#### Scenario: Correct physical size at 300 DPI

- **WHEN** a mermaid PNG is 1200px wide at 300 DPI
- **THEN** the embedded image width SHALL be 4 inches (1200/300 * 914400 = 3657600 EMU)

### Requirement: Mermaid configuration additions

The system SHALL support the following additional fields in the `[mermaid]` section of `md2docx.toml`:

- `output_format`: `"png"` (default) or `"svg"`
- `dpi`: PNG resolution as integer (default: `150`)

#### Scenario: Configuration section parsed with new fields

- **WHEN** `md2docx.toml` contains `[mermaid]` with `output_format = "png"` and `dpi = 200`
- **THEN** the system SHALL parse both fields and use them for mermaid rendering

#### Scenario: Partial configuration

- **WHEN** `md2docx.toml` contains `[mermaid]` with only `dpi = 300` (no `output_format`)
- **THEN** the system SHALL default `output_format` to `"png"` and use 300 DPI

### Requirement: Fallback when PNG feature unavailable

The system SHALL fall back to SVG rendering when the `mermaid-png` compile-time feature is not enabled, regardless of the configured `output_format`.

#### Scenario: Feature not compiled in

- **WHEN** `output_format = "png"` but the binary was compiled without the `mermaid-png` feature
- **THEN** the system SHALL fall back to SVG rendering silently

### Requirement: PNG filename in DOCX package

The system SHALL use `.png` file extension for mermaid image filenames in the DOCX package when outputting PNG, ensuring correct OOXML content type registration.

#### Scenario: PNG filename

- **WHEN** a mermaid diagram is rendered as PNG
- **THEN** the virtual filename SHALL use `.png` extension (e.g., `mermaid1.png`)

#### Scenario: SVG filename preserved

- **WHEN** a mermaid diagram is rendered as SVG
- **THEN** the virtual filename SHALL use `.svg` extension (e.g., `mermaid1.svg`)
