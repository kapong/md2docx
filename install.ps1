$ErrorActionPreference = "Stop"

$Repo = "kapong/md2docx"
$InstallDir = "$env:USERPROFILE\.md2docx\bin"

# Get latest release version
function Get-LatestVersion {
    Write-Host "Checking for latest release..."
    try {
        $response = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest"
        return $response.tag_name
    } catch {
        Write-Error "Failed to get latest release: $_"
        exit 1
    }
}

# Download file
function Download-File {
    param(
        [string]$Url,
        [string]$OutputPath
    )
    
    Write-Host "Downloading from $Url..."
    try {
        Invoke-WebRequest -Uri $Url -OutFile $OutputPath
    } catch {
        Write-Error "Failed to download file: $_"
        exit 1
    }
}

# Verify SHA256 checksum
function Verify-Checksum {
    param(
        [string]$FilePath,
        [string]$ChecksumsPath,
        [string]$BinaryName
    )
    
    Write-Host "Verifying checksum..."
    
    # Read checksums file
    $checksums = Get-Content $ChecksumsPath
    
    # Find the line for our binary
    $checksumLine = $checksums | Where-Object { $_ -match "$BinaryName" }
    
    if (-not $checksumLine) {
        Write-Error "Checksum not found for $BinaryName"
        exit 1
    }
    
    # Extract expected hash
    $expectedHash = ($checksumLine -split '\s+')[0]
    
    # Calculate actual hash
    $actualHash = (Get-FileHash -Path $FilePath -Algorithm SHA256).Hash.ToLower()
    
    if ($actualHash -ne $expectedHash) {
        Write-Error "Checksum verification failed!"
        Write-Error "Expected: $expectedHash"
        Write-Error "Got:      $actualHash"
        exit 1
    }
    
    Write-Host "✓ Checksum verified"
}

# Add to PATH if not already present
function Add-ToPath {
    param(
        [string]$Directory
    )
    
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
    
    if ($currentPath -notlike "*$Directory*") {
        Write-Host "Adding $Directory to PATH..."
        [Environment]::SetEnvironmentVariable(
            "Path",
            "$currentPath;$Directory",
            "User"
        )
        Write-Host "✓ Added to PATH"
        return $true
    } else {
        Write-Host "✓ Already in PATH"
        return $false
    }
}

# Main installation
function Install-Md2docx {
    Write-Host "md2docx installer"
    Write-Host "================"
    Write-Host ""
    
    $version = Get-LatestVersion
    
    if (-not $version) {
        Write-Error "Could not determine latest version"
        exit 1
    }
    
    Write-Host "Installing md2docx $version..."
    Write-Host ""
    
    # Create temp directory
    $tempDir = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "md2docx-install-$(Get-Random)")
    
    try {
        # Download binary
        $binaryName = "md2docx-windows-x86_64.exe"
        $binaryPath = Join-Path $tempDir $binaryName
        $downloadUrl = "https://github.com/$Repo/releases/download/$version/$binaryName"
        Download-File -Url $downloadUrl -OutputPath $binaryPath
        
        # Download checksums
        $checksumsPath = Join-Path $tempDir "checksums.txt"
        $checksumsUrl = "https://github.com/$Repo/releases/download/$version/checksums.txt"
        Download-File -Url $checksumsUrl -OutputPath $checksumsPath
        
        # Verify checksum
        Verify-Checksum -FilePath $binaryPath -ChecksumsPath $checksumsPath -BinaryName $binaryName
        
        # Create install directory
        if (-not (Test-Path $InstallDir)) {
            Write-Host "Creating install directory..."
            New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        }
        
        # Move binary to install directory
        $finalPath = Join-Path $InstallDir "md2docx.exe"
        Write-Host "Installing to $finalPath..."
        Move-Item -Path $binaryPath -Destination $finalPath -Force
        
        # Add to PATH
        $pathModified = Add-ToPath -Directory $InstallDir
        
        Write-Host ""
        Write-Host "✓ md2docx $version installed successfully!"
        Write-Host ""
        
        if ($pathModified) {
            Write-Host "⚠ Please restart your terminal for PATH changes to take effect."
            Write-Host ""
        }
        
        Write-Host "Run 'md2docx --help' to get started."
        
    } finally {
        # Cleanup temp directory
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force
        }
    }
}

# Run installation
Install-Md2docx
