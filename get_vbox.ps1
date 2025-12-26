# VBox-PowerPortable
# A PowerShell script to get the latest version of VirtualBox and extract it to a portable folder.
# Author: CypherpunkSamurai (github.com)
# Date: 2024-06-17

# Header
$VERSION = "1.0"
$AUTHOR = "CypherpunkSamurai"
$HEADER = "        __  ___                 __   __          __   __   __  ___       __        ___    
\  / | |__)  |  |  |  /\  |    |__) /  \ \_/    |__) /  \ |__)  |   /\  |__) |    |__     
 \/  | |  \  |  \__/ /~~\ |___ |__) \__/ / \    |    \__/ |  \  |  /~~\ |__) |___ |___    
                                                                                          "
Write-Host "`n$HEADER"
Write-Host "VBox-PowerPortable v$VERSION - by $AUTHOR`n" -ForegroundColor Green
Write-Host "   ~ Getter Script ~`n" -ForegroundColor Green


# Constants
$VBOX_WEB_URL = "https://www.virtualbox.org/wiki/Downloads"
$CACHE_FOLDER = "$PSScriptRoot\.cache"

# Variables
$VERBOSE = $false
# check arguments
if ($args -contains "-v" -or $args -contains "--verbose" -or $args -contains "-verbose" -or $args -contains "--v") {
    Write-Host "[i] Verbose mode enabled" -ForegroundColor Green
    $VERBOSE = $true
}
$TARGET_FOLDER = Read-Host "[>] Where to install VirtualBox (default: [CURRENT_FOLDER]\vbox) ?"
if ($TARGET_FOLDER -eq "") {
    $TARGET_FOLDER = "$PSScriptRoot\vbox"
} else {
    $TARGET_FOLDER = "$TARGET_FOLDER\vbox"
}

# Logging Functions
function Log-Info {
    param ([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Log-Verbose {
    param ([string]$Message)
    if ($VERBOSE) {
        Write-Host "[VERBOSE] $Message" -ForegroundColor Gray
    }
}

function Log-Success {
    param ([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Log-Warning {
    param ([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Log-Error {
    param ([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function Log-Step {
    param ([string]$Message)
    Write-Host "`n=== $Message ===" -ForegroundColor Magenta
}

# Print configuration
Log-Info "Virtualbox Target Folder: $TARGET_FOLDER"
Log-Info "Cache Folder: $CACHE_FOLDER"

# Functions
function Make-Temp-Dir {
    $TEMP_DIR = [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
    Log-Verbose "Creating temporary directory: $TEMP_DIR"
    New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null
    Log-Verbose "Temporary directory created successfully"
    return $TEMP_DIR
} # Make-Temp-Dir function

function Ensure-Cache-Dir {
    if (-not (Test-Path $CACHE_FOLDER)) {
        Log-Info "Creating cache folder: $CACHE_FOLDER"
        New-Item -ItemType Directory -Path $CACHE_FOLDER -Force | Out-Null
        Log-Success "Cache folder created"
    } else {
        Log-Verbose "Cache folder already exists: $CACHE_FOLDER"
    }
} # Ensure-Cache-Dir function

function Get-VBox-URL {
    Log-Step "Fetching VirtualBox Download URL"
    Log-Verbose "Querying VirtualBox downloads page: $VBOX_WEB_URL"
    
    try {
        $web_response = Invoke-WebRequest -Uri $VBOX_WEB_URL -UseBasicParsing
        Log-Verbose "Web response received, status: $($web_response.StatusCode)"
        Log-Verbose "Parsing page for download links..."
        
        $VBOX_URL = $web_response.Links | Where-Object { $_.href -match "VirtualBox-.*-Win\.exe$" } | ForEach-Object { $_.href } | Select-Object -First 1
        
        if (-not $VBOX_URL) {
            Log-Warning "Could not find Windows VirtualBox URL via regex, trying class-based lookup..."
            $VBOX_URL = $web_response.Links | Where-Object { $_.class -eq "ext-link" } | ForEach-Object { $_.href } | Select-Object -First 1
        }
        
        if ($VBOX_URL) {
            Log-Success "Found VirtualBox download URL"
            Log-Info "URL: $VBOX_URL"
        } else {
            Log-Error "Could not find VirtualBox download URL"
            throw "VirtualBox URL not found"
        }
        
        return $VBOX_URL
    } catch {
        Log-Error "Failed to fetch VirtualBox URL: $($_.Exception.Message)"
        throw
    }
} # Get-VBox-URL function

function Get-Cached-Installer {
    param ([string]$URL)
    
    Log-Step "Checking for Cached Installer"
    Ensure-Cache-Dir
    
    # Extract filename from URL
    $FileName = [System.IO.Path]::GetFileName($URL)
    $CachedFile = "$CACHE_FOLDER\$FileName"
    
    Log-Verbose "Expected cache file: $CachedFile"
    
    if (Test-Path $CachedFile) {
        $FileInfo = Get-Item $CachedFile
        $FileSizeMB = [math]::Round($FileInfo.Length / 1MB, 2)
        $FileDate = $FileInfo.LastWriteTime
        
        Log-Success "Found cached installer!"
        Log-Info "  File: $FileName"
        Log-Info "  Size: $FileSizeMB MB"
        Log-Info "  Last Modified: $FileDate"
        
        return $CachedFile
    }
    
    Log-Info "No cached installer found for: $FileName"
    return $null
} # Get-Cached-Installer function

function Get-File-From-URL {
    param (
        [string]$URL,
        [string]$TARGET_FILE,
        [switch]$UseCache
    )
    
    Log-Step "Downloading File"
    Log-Info "Source URL: $URL"
    Log-Info "Target File: $TARGET_FILE"
    
    # check if the file already exists
    if (Test-Path $TARGET_FILE) {
        Log-Warning "File already exists at target: $TARGET_FILE"
        throw "File already exists"
        return
    }
    
    Log-Verbose "Starting download..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    try {
        # Use BITS for better progress and resume support
        Log-Verbose "Using Invoke-WebRequest for download..."
        $ProgressPreference = 'Continue'
        Invoke-WebRequest -Uri $URL -OutFile $TARGET_FILE -UseBasicParsing
        
        $stopwatch.Stop()
        $elapsed = $stopwatch.Elapsed
        
        if (Test-Path $TARGET_FILE) {
            $FileInfo = Get-Item $TARGET_FILE
            $FileSizeMB = [math]::Round($FileInfo.Length / 1MB, 2)
            $SpeedMBps = [math]::Round($FileSizeMB / $elapsed.TotalSeconds, 2)
            
            Log-Success "Download completed!"
            Log-Info "  Size: $FileSizeMB MB"
            Log-Info "  Time: $($elapsed.ToString('mm\:ss'))"
            Log-Info "  Speed: $SpeedMBps MB/s"
        }
    } catch {
        Log-Error "Download failed: $($_.Exception.Message)"
        throw
    }
    
    return $true
} # Get-File-From-URL function

function Download-Or-Use-Cached {
    param ([string]$URL)
    
    # Check cache first
    $CachedFile = Get-Cached-Installer -URL $URL
    
    if ($CachedFile) {
        Log-Success "Using cached installer: $CachedFile"
        return $CachedFile
    }
    
    # Download to cache
    $FileName = [System.IO.Path]::GetFileName($URL)
    $CacheTarget = "$CACHE_FOLDER\$FileName"
    
    Log-Info "Downloading installer to cache..."
    Get-File-From-URL -URL $URL -TARGET_FILE $CacheTarget
    
    Log-Success "Installer cached for future use"
    return $CacheTarget
} # Download-Or-Use-Cached function

function Get-Msi-From-Exe {
    param (
        [string]$EXE_FILE,
        [string]$TARGET_FOLDER,
        [string]$TARGET_FILENAME
    )
    
    Log-Step "Extracting MSI from EXE"
    Log-Info "Source EXE: $EXE_FILE"
    Log-Info "Target Folder: $TARGET_FOLDER"
    Log-Info "Target Filename: $TARGET_FILENAME"
    
    # check if the file exists
    if (-not (Test-Path $EXE_FILE)) {
        Log-Error "Source EXE not found: $EXE_FILE"
        throw "File not found"
        return
    }
    Log-Verbose "Source EXE exists: OK"
    
    # check if the folder exists
    if (-not (Test-Path $TARGET_FOLDER)) {
        Log-Error "Target folder not found: $TARGET_FOLDER"
        throw "Folder not found"
        return
    }
    Log-Verbose "Target folder exists: OK"
    
    # check if the file already exists
    if (Test-Path "$TARGET_FOLDER\$TARGET_FILENAME") {
        Log-Warning "Target MSI already exists: $TARGET_FOLDER\$TARGET_FILENAME"
        throw "File already exists"
        return
    }

    # extract the MSI file
    Log-Verbose "Creating temporary extraction directory..."
    $FUNCTION_TEMP_DIR = Make-Temp-Dir
    
    Log-Info "Running VirtualBox installer in extraction mode..."
    Log-Verbose "Command: cmd /min /c set __COMPAT_LAYER=RUNASINVOKER && start `"`" `"$EXE_FILE`" --silent -extract -path `"$FUNCTION_TEMP_DIR`""
    
    $extractProcess = Start-Process cmd -ArgumentList "/min /c set __COMPAT_LAYER=RUNASINVOKER && start `"`" `"$EXE_FILE`" --silent -extract -path `"$FUNCTION_TEMP_DIR`"" -Wait -NoNewWindow -PassThru
    
    Log-Verbose "Extraction process completed with exit code: $($extractProcess.ExitCode)"
    
    # Wait a bit for files to be written
    Start-Sleep -Seconds 2
    
    # Find extracted MSI
    Log-Verbose "Looking for extracted MSI files in: $FUNCTION_TEMP_DIR"
    $MsiFiles = Get-ChildItem -Path $FUNCTION_TEMP_DIR -Filter "VirtualBox-*.msi" -ErrorAction SilentlyContinue
    
    if ($MsiFiles) {
        foreach ($msi in $MsiFiles) {
            Log-Verbose "  Found: $($msi.Name)"
        }
        
        Move-Item $FUNCTION_TEMP_DIR\VirtualBox-*.msi -Destination "$TARGET_FOLDER\$TARGET_FILENAME" -Force -ErrorAction SilentlyContinue | Out-Null
        Log-Success "MSI extracted and moved: $TARGET_FOLDER\$TARGET_FILENAME"
    } else {
        Log-Warning "No MSI files found in extraction directory"
        Log-Verbose "Directory contents:"
        Get-ChildItem -Path $FUNCTION_TEMP_DIR -Recurse | ForEach-Object { Log-Verbose "  $($_.FullName)" }
    }
    
    Log-Verbose "Cleaning up temporary extraction directory..."
    Remove-Item -Path $FUNCTION_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    Log-Verbose "Cleanup completed"
}

function Extract-Msi {
    param (
        [string]$MSI_FILE,
        [string]$TARGET_FOLDER
    )
    
    Log-Step "Extracting MSI Contents"
    Log-Info "Source MSI: $MSI_FILE"
    Log-Info "Target Folder: $TARGET_FOLDER"
    
    # check if the file exists
    if (-not (Test-Path $MSI_FILE)) {
        Log-Error "MSI file not found: $MSI_FILE"
        throw "File not found"
        return
    }
    Log-Verbose "MSI file exists: OK"
    
    # check if the folder exists
    if (-not (Test-Path $TARGET_FOLDER)) {
        Log-Error "Target folder not found: $TARGET_FOLDER"
        throw "Folder not found"
        return
    }
    Log-Verbose "Target folder exists: OK"

    # Extract the MSI file
    $FUNCTION_TEMP_DIR = Make-Temp-Dir
    
    Log-Info "Extracting MSI to temp folder..."
    $msiCmd = "msiexec /a `"$MSI_FILE`" /qn TARGETDIR=`"$FUNCTION_TEMP_DIR`""
    Log-Verbose "Running command: $msiCmd"
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    cmd /c $msiCmd
    $stopwatch.Stop()
    
    Log-Verbose "MSI extraction completed in: $($stopwatch.Elapsed.ToString('mm\:ss'))"
    
    # Find VirtualBox.exe in the extracted files
    Log-Verbose "Searching for VirtualBox.exe in extracted files..."
    $VBoxExe = Get-ChildItem -Path $FUNCTION_TEMP_DIR -Recurse -Filter "VirtualBox.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($VBoxExe) {
        $VBoxDir = $VBoxExe.DirectoryName
        Log-Success "Found VirtualBox at: $VBoxDir"
        
        # Count files to copy
        $FilesToCopy = Get-ChildItem -Path $VBoxDir -Recurse -File
        $FileCount = $FilesToCopy.Count
        $TotalSizeMB = [math]::Round(($FilesToCopy | Measure-Object -Property Length -Sum).Sum / 1MB, 2)
        
        Log-Info "Copying $FileCount files ($TotalSizeMB MB) to target..."
        Log-Verbose "Source: $VBoxDir"
        Log-Verbose "Target: $TARGET_FOLDER"
        
        Copy-Item "$VBoxDir\*" $TARGET_FOLDER -Recurse -Force
        
        Log-Success "VirtualBox extracted successfully!"
        
        # Verify installation
        if (Test-Path "$TARGET_FOLDER\VirtualBox.exe") {
            Log-Success "Verification: VirtualBox.exe found in target folder"
        } else {
            Log-Warning "Verification: VirtualBox.exe NOT found in target folder!"
        }
    } else {
        Log-Error "VirtualBox.exe not found in extracted MSI!"
        Log-Info "Temp directory contents:"
        Get-ChildItem -Path $FUNCTION_TEMP_DIR -Recurse -Directory | ForEach-Object { Log-Verbose "  $($_.FullName)" }
    }
    
    Log-Verbose "Cleaning up temporary extraction directory..."
    Remove-Item -Path $FUNCTION_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    Log-Verbose "Cleanup completed"
}

# System Checks
Log-Step "System Checks"

Log-Verbose "Checking for msiexec.exe..."
if (-not (Test-Path "C:\Windows\System32\msiexec.exe")) {
    Log-Error "msiexec.exe not found. Please make sure you have Windows Installer installed."
    exit 1
}
Log-Success "msiexec.exe found: OK"

Log-Verbose "Checking if target folder already exists..."
if (Test-Path "$TARGET_FOLDER") {
    Log-Error "Target folder already exists: $TARGET_FOLDER"
    Log-Info "Please remove the existing folder or choose a different location."
    exit 1
}
Log-Success "Target folder is available: OK"

# Main
Log-Step "Starting VirtualBox Installation"
Log-Info "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$TEMP_DIR = Make-Temp-Dir
try {
    # Get the latest version of VirtualBox
    $VBOX_URL = Get-VBox-URL
    
    # Download or use cached installer
    $VBOX_FILE = Download-Or-Use-Cached -URL $VBOX_URL

    # Extract the EXE file
    Get-Msi-From-Exe -EXE_FILE $VBOX_FILE -TARGET_FOLDER $TEMP_DIR -TARGET_FILENAME "virtualbox.msi"
    Log-Verbose "VirtualBox MSI file extracted successfully."

    # Check if the target folder exists
    if (-not (Test-Path $TARGET_FOLDER)) {
        Log-Info "Creating target folder: $TARGET_FOLDER"
        New-Item -ItemType Directory -Path $TARGET_FOLDER -Force | Out-Null
        Log-Verbose "Target folder created"
    }

    # Extract the MSI file
    Extract-Msi -MSI_FILE "$TEMP_DIR\virtualbox.msi" -TARGET_FOLDER $TARGET_FOLDER

    # Success
    Log-Step "Installation Complete"
    Log-Success "VirtualBox is now located at: $TARGET_FOLDER"
    Log-Info "Cached installer available at: $CACHE_FOLDER"
    Log-Info "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Log-Success "Completed successfully!"
} catch {
    Log-Error "Failed to get VirtualBox."
    Log-Error $_.Exception.Message
    Log-Verbose "Stack Trace: $($_.ScriptStackTrace)"
    exit 1
} finally {
    # cleanup
    Log-Verbose "Cleaning up temporary files: $TEMP_DIR"
    Remove-Item -Path $TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue
    Log-Verbose "Final cleanup completed"
}