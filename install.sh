#!/bin/bash
set -e

REPO="kapong/md2docx"
INSTALL_DIR="${INSTALL_DIR:-/usr/local/bin}"

# Check for Windows environments and redirect to PowerShell installer
case "$(uname -s)" in
    MINGW*|MSYS*|CYGWIN*)
        echo "Windows detected. Please use the PowerShell installer instead:"
        echo "  irm https://raw.githubusercontent.com/kapong/md2docx/main/install.ps1 | iex"
        exit 1
        ;;
esac

# Detect OS and architecture
detect_platform() {
    local os arch
    
    case "$(uname -s)" in
        Darwin) os="darwin" ;;
        Linux)  os="linux" ;;
        *)      echo "Unsupported OS: $(uname -s)" && exit 1 ;;
    esac
    
    case "$(uname -m)" in
        x86_64|amd64)   arch="x86_64" ;;
        arm64|aarch64)  arch="arm64" ;;
        *)              echo "Unsupported architecture: $(uname -m)" && exit 1 ;;
    esac
    
    # Linux only supports x86_64
    if [ "$os" = "linux" ] && [ "$arch" = "arm64" ]; then
        echo "Linux arm64 is not supported yet"
        exit 1
    fi
    
    echo "${os}-${arch}"
}

# Get latest release version
get_latest_version() {
    curl -fsSL "https://api.github.com/repos/${REPO}/releases/latest" | grep '"tag_name"' | sed -E 's/.*"([^"]+)".*/\1/'
}

# Download and verify binary
download_binary() {
    local platform="$1"
    local version="$2"
    local binary_name="md2docx-${platform}"
    local download_url="https://github.com/${REPO}/releases/download/${version}/${binary_name}"
    local checksums_url="https://github.com/${REPO}/releases/download/${version}/checksums.txt"
    
    local tmp_dir
    tmp_dir=$(mktemp -d)
    trap 'rm -rf "$tmp_dir"' EXIT
    
    echo "Downloading md2docx ${version} for ${platform}..."
    curl -fsSL -o "${tmp_dir}/md2docx" "$download_url"
    
    echo "Downloading checksums..."
    curl -fsSL -o "${tmp_dir}/checksums.txt" "$checksums_url"
    
    echo "Verifying checksum..."
    cd "$tmp_dir"
    if command -v sha256sum &> /dev/null; then
        grep "${binary_name}" checksums.txt | sed "s/${binary_name}/md2docx/" | sha256sum -c -
    elif command -v shasum &> /dev/null; then
        grep "${binary_name}" checksums.txt | sed "s/${binary_name}/md2docx/" | shasum -a 256 -c -
    else
        echo "Warning: Cannot verify checksum (no sha256sum or shasum found)"
    fi
    
    chmod +x "${tmp_dir}/md2docx"
    
    echo "Installing to ${INSTALL_DIR}/md2docx..."
    if [ -w "$INSTALL_DIR" ]; then
        mv "${tmp_dir}/md2docx" "${INSTALL_DIR}/md2docx"
    else
        sudo mv "${tmp_dir}/md2docx" "${INSTALL_DIR}/md2docx"
    fi
}

main() {
    echo "md2docx installer"
    echo "================"
    
    local platform version
    platform=$(detect_platform)
    version=$(get_latest_version)
    
    if [ -z "$version" ]; then
        echo "Error: Could not determine latest version"
        exit 1
    fi
    
    download_binary "$platform" "$version"
    
    echo ""
    echo "✓ md2docx ${version} installed successfully!"
    echo ""
    echo "Run 'md2docx --help' to get started."
}

main "$@"
