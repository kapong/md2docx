# md2docx Development Plan

## Tomorrow's Tasks

### 1. Code Review & Refactoring (High Priority)

#### Error Handling Improvements
- [x] Replace remaining `unwrap()` calls with proper error handling using `?` or `expect()` with descriptive messages
  - Status: Completed - mermaid/mod.rs regex unwraps converted to `once_cell::sync::Lazy` static regexes
- [x] Review error types in `src/error.rs` and ensure comprehensive error coverage
- [ ] Add context to errors where appropriate using `anyhow` or similar

#### Code Quality (Address Clippy Warnings)
- [x] Simplify complex types in `src/docx/builder.rs` (headers/footers Vec types)
  - Created `HeaderFooterEntry` and `MediaFileMapping` structs to replace complex tuple types
- [x] Refactor functions with too many arguments:
  - [x] `BuildContext::new()` (15 args)
  - [x] `create_table_cell_with_template()` (8 args)
  - [x] `Paragraph::with_page_layout()` (10 args)
  - [x] `apply_cover_template()` (8 args) - Created `CoverTemplateContext` struct
- [x] Collapse nested `if let` statements where possible
  - Refactored 3 locations in builder.rs using `Option::and_then`
- [ ] Convert loops to `while let` where appropriate (low priority - current loops are idiomatic)

#### Documentation
- [ ] Add missing doc comments to public APIs
- [ ] Include examples in doc comments
- [ ] Review and update module-level documentation

#### Code Organization
- [x] Review module visibility (pub vs pub(crate) vs private)
  - Made internal types `pub(crate)`: ImageContext, HyperlinkContext, NumberingContext, BuildResult, BuildContext, etc.
  - OOXML types now properly encapsulated (HeaderFooterRefs, BookmarkStart, DocumentXml, etc.)
  - Template extraction types visibility cleaned up
- [x] Check for code duplication and extract common functionality
  - Created `src/template/extract/xml_utils.rs` with shared utilities
  - Extracted `extract_attribute()` function (was 3 copies)
  - Extracted `extract_run_properties()` with configurable defaults (was 3 copies)
  - ~100+ lines of duplicate code removed
- [ ] Ensure consistent naming conventions

### 2. Docker Setup (High Priority) - **Antigravity Mission**

> **Note**: This task is handled by Antigravity, not OpenCode.

#### Dockerfile (Ultimate: Alpine for minimal size)
```dockerfile
# syntax=docker/dockerfile:1
# Multi-stage build: Rust → Alpine (ultimate minimal image)

# Stage 1: Build with latest Rust toolchain
FROM rust:1-alpine AS builder
WORKDIR /app

# Install build dependencies
RUN apk add --no-cache musl-dev openssl-dev pkgconfig

# Copy manifests first for better caching
COPY Cargo.toml Cargo.lock ./
COPY src ./src

# Build static binary
ENV RUSTFLAGS="-C target-feature=+crt-static"
RUN cargo build --release --target x86_64-unknown-linux-musl

# Stage 2: Ultimate minimal runtime with Alpine
FROM alpine:latest
RUN apk add --no-cache ca-certificates

# Create non-root user
RUN adduser -D -u 1000 md2docx

# Copy static binary from builder
COPY --from=builder --chown=md2docx:md2docx \
    /app/target/x86_64-unknown-linux-musl/release/md2docx \
    /usr/local/bin/

# Switch to non-root user
USER md2docx
WORKDIR /workspace

# Set entrypoint
ENTRYPOINT ["md2docx"]
CMD ["--help"]
```

#### Alternative: Debian Trixie (if Alpine has compatibility issues)
```dockerfile
# syntax=docker/dockerfile:1
# Multi-stage build with latest Rust and Debian Trixie

# Stage 1: Build with latest Rust toolchain
FROM rust:1-slim-bookworm AS builder
WORKDIR /app

# Install dependencies for building
RUN apt-get update && apt-get install -y \
    pkg-config \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy manifests first for better caching
COPY Cargo.toml Cargo.lock ./
COPY src ./src

# Build release binary
RUN cargo build --release

# Stage 2: Minimal runtime image with Debian Trixie
FROM debian:trixie-slim
RUN apt-get update && apt-get install -y \
    ca-certificates \
    libssl3 \
    && rm -rf /var/lib/apt/lists/*

# Create non-root user
RUN useradd --create-home --uid 1000 --shell /bin/bash md2docx

# Copy binary from builder
COPY --from=builder --chown=md2docx:md2docx /app/target/release/md2docx /usr/local/bin/

# Switch to non-root user
USER md2docx
WORKDIR /workspace

# Set entrypoint
ENTRYPOINT ["md2docx"]
CMD ["--help"]
```

#### compose.yaml (Docker Compose v2+ syntax)
```yaml
name: md2docx

services:
  md2docx:
    build:
      context: .
      dockerfile: Dockerfile
      target: runtime
    image: md2docx:latest
    container_name: md2docx
    volumes:
      - type: bind
        source: ./docs
        target: /workspace/docs
        read_only: true
      - type: bind
        source: ./output
        target: /workspace/output
    working_dir: /workspace
    command: build -d /workspace/docs -o /workspace/output/output.docx
    # Or use with custom config:
    # command: build -d /workspace/docs --config /workspace/docs/md2docx.toml
    
  # Alternative service for single file conversion
  convert:
    extends:
      service: md2docx
    profiles: ["convert"]
    command: build -i /workspace/docs/input.md -o /workspace/output/output.docx
```

#### .dockerignore
```
# Git
.git
.gitignore

# Rust build artifacts
target/
Cargo.lock

# IDE
.idea/
.vscode/
*.swp
*.swo

# OS
.DS_Store
Thumbs.db

# Generated files
*.docx
output/
examples/**/output/

# Documentation (keep for build context)
!README.md
!LICENSE

# Temp files
~$*
*.tmp
```

#### Tasks
- [ ] Create optimized Dockerfile with multi-stage build (Alpine ultimate target)
- [ ] Create compose.yaml with Docker Compose v2+ syntax
- [ ] Add .dockerignore file
- [ ] Test Docker build and run (both Alpine and Debian variants)
- [ ] Compare image sizes: Alpine vs Debian Trixie
- [ ] Document Docker usage in README

#### Docker Image Size Comparison
| Base Image | Expected Size | Use Case |
|------------|---------------|----------|
| Alpine | ~15-20 MB | Production, minimal footprint |
| Debian Trixie | ~50-80 MB | Compatibility, debugging |
| Debian Bookworm | ~60-90 MB | Stable fallback |

**Recommendation**: Use Alpine as default, provide Debian variant for troubleshooting

### 3. Git Rebase for Public Release (High Priority)

#### History Cleanup
- [ ] Review git log: `git log --oneline --graph`
- [ ] Identify commits to squash:
  - Debug commits
  - Fix-up commits
  - WIP commits
- [ ] Interactive rebase: `git rebase -i HEAD~N`
- [ ] Rewrite commit messages to follow conventional commits:
  - `feat:` - New features
  - `fix:` - Bug fixes
  - `refactor:` - Code refactoring
  - `docs:` - Documentation
  - `chore:` - Maintenance

#### Security & Cleanup
- [x] Scan for sensitive data:
  - [x] Hardcoded paths (check for `/Users/kapong/`) - Fixed 2 test files to use `env!("CARGO_MANIFEST_DIR")`
  - [x] API keys or tokens - None found
  - [x] Personal information - None found (test emails use example.com)
- [x] Update .gitignore:
  - [x] `*.docx` already present, added exception for template files
  - [x] `output/` directories already present
  - [x] `~$*.docx` temp files already present
  - [x] Added `plan.md` exclusion for public release
- [ ] Clean working directory

#### Release Preparation
- [x] Create `LICENSE` file (MIT) - Already exists
- [x] Create `CHANGELOG.md` with version history
- [x] Create `CONTRIBUTING.md` with guidelines
- [ ] Tag initial release: `git tag -a v0.1.0 -m "Initial release"`

### 4. CI/CD Pipeline (GitHub Actions)

#### Workflow: Build and Push Docker Image (`.github/workflows/docker-publish.yml`)
```yaml
name: Build and Push Docker Image

on:
  push:
    branches: [ "main" ]
    # Publish semver tags as releases.
    tags: [ "v*.*.*" ]
  pull_request:
    branches: [ "main" ]

env:
  REGISTRY: ghcr.io
  IMAGE_NAME: ${{ github.repository }}

jobs:
  build:
    runs-on: ubuntu-latest
    permissions:
      contents: read
      packages: write
      # This is used to complete the identity challenge
      # with sigstore/fulcio when running outside of PRs.
      id-token: write

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      # Set up Buildx for multi-platform builds (if needed in future)
      - name: Set up Docker Buildx
        uses: docker/setup-buildx-action@v3

      # Login to GHCR
      - name: Log into registry ${{ env.REGISTRY }}
        if: github.event_name != 'pull_request'
        uses: docker/login-action@v3
        with:
          registry: ${{ env.REGISTRY }}
          username: ${{ github.actor }}
          password: ${{ secrets.GITHUB_TOKEN }}

      # Extract metadata (tags, labels) for Docker
      - name: Extract Docker metadata
        id: meta
        uses: docker/metadata-action@v5
        with:
          images: ${{ env.REGISTRY }}/${{ env.IMAGE_NAME }}
          tags: |
            type=raw,value=latest,enable={{is_default_branch}}
            type=semver,pattern={{version}}
            type=sha,format=long

      # Build and push Docker image
      - name: Build and push Docker image
        id: build-and-push
        uses: docker/build-push-action@v5
        with:
          context: .
          push: ${{ github.event_name != 'pull_request' }}
          tags: ${{ steps.meta.outputs.tags }}
          labels: ${{ steps.meta.outputs.labels }}
          platforms: linux/amd64
          cache-from: type=gha
          cache-to: type=gha,mode=max
```

#### Tasks
- [ ] Create `.github/workflows/docker-publish.yml`
- [ ] Configure repository permissions (Package settings: Allow write access)
- [ ] Test workflow with a push to main
- [ ] Verify image exists in GitHub Package Registry (GHCR)

### 5. Documentation Updates (Medium Priority)

#### README.md Updates
- [ ] Add Docker usage section
- [ ] Add installation instructions
- [ ] Add quick start guide
- [ ] Add configuration examples
- [ ] Add troubleshooting section

#### Additional Documentation
- [ ] Create `docs/ARCHITECTURE.md` explaining codebase structure
- [ ] Create `docs/API.md` for library usage
- [ ] Create `docs/DOCKER.md` for Docker-specific instructions

## Implementation Order

1. **Start with Git cleanup** - Do this first to have clean history
2. **Code refactoring** - Address clippy warnings and improve quality
3. **Docker setup** - Create containerization
4. **Documentation** - Update all docs last when code is stable

## Notes

- Keep commits atomic and focused
- Test thoroughly after each major change
- Ensure backward compatibility if possible
- Consider adding CI/CD workflow for automated testing
- Think about publishing to crates.io after public release
