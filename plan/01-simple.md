# Simplified Mermaid Support Plan

## Overview
This plan outlines a simplified approach to Mermaid diagram support in `md2docx` by leveraging the external `mmdc` (Mermaid CLI) tool instead of embedding a headless browser.

## Comparison: Complex vs. Simple

| Feature | Complex Approach | Simple Approach (This Plan) |
|---------|------------------|-----------------------------|
| **Core Engine** | `chromiumoxide` (Headless Chrome) | `mmdc` (External CLI) |
| **Concurrency** | `tokio` (Async/await) | `std::process::Command` (Sync) |
| **Dependencies** | High (tokio, chromiumoxide, etc.) | Low (None extra) |
| **Binary Size** | Large (Bundles/downloads Chrome) | Small |
| **Complexity** | High (Managing browser lifecycle) | Low (Process execution) |
| **Ease of Use** | High (Batteries included) | Medium (Requires NPM install) |

## Module Structure

The Mermaid functionality will be contained within `src/mermaid/`:

- `src/mermaid/mod.rs`: Main entry point. Handles the rendering logic, process management, and fallback.
- `src/mermaid/cache.rs`: Hash-based caching logic (using SHA-256) to avoid re-rendering unchanged diagrams.
- `src/mermaid/config.rs`: Configuration struct for Mermaid (theme, background, width, scale).

## Implementation Steps

1. **Define Config**: Create `MermaidConfig` in `config.rs`.
2. **Implement Caching**:
    - Hash the Mermaid code block content.
    - Check for existing `.png` in `.md2docx-cache/mermaid/`.
3. **Execute `mmdc`**:
    - Check if `mmdc` is available in `PATH`.
    - Write Mermaid content to a temporary file.
    - Run `mmdc -i temp.mmd -o output.png`.
    - Read output bytes.
4. **Graceful Fallback**:
    - If `mmdc` is not found or fails, return the original Mermaid source as a standard markdown code block.
5. **Integration**:
    - Hook into `src/docx/builder.rs` to replace Mermaid code blocks with images.

## Installation Requirements

Users must install the Mermaid CLI tool via NPM:

```bash
npm install -g @mermaid-js/mermaid-cli
```

## Error Handling

- **Missing `mmdc`**: Log a warning and fallback to rendering the Mermaid source as a text code block.
- **Render Failure**: Log the error from `mmdc` and fallback to text.
- **Cache Failures**: If cache writing fails, proceed without caching (warn user).

## Pros & Cons

### Pros
- **Fast Development**: Extremely quick to implement using `std::process`.
- **Minimal Bloat**: Keeps the `md2docx` binary small and dependency-free.
- **Reliability**: `mmdc` is the official CLI and handles complex diagrams and themes natively.

### Cons
- **External Dependency**: Requires Node.js and `mmdc` to be installed on the user's system.
- **Process Overhead**: Spawning a process for each diagram (mitigated by caching).
