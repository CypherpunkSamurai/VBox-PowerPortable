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
Write-Host "   ~ Start Script ~`n" -ForegroundColor Green

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

# Elevate the script
Log-Info "Checking for administrator privileges..."
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Log-Warning "Not running as Administrator. Elevating..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Log-Success "Running with Administrator privileges"

# Variables
$VERBOSE = $false
# check arguments
if ($args -contains "-v" -or $args -contains "--verbose" -or $args -contains "-verbose" -or $args -contains "--v") {
    Log-Info "Verbose mode enabled"
    $VERBOSE = $true
}

$TOOLS_FOLDER = "$PSScriptRoot\tools"
$TARGET_FOLDER = Read-Host "[>] Where to setup VirtualBox (default: [CURRENT_FOLDER]\vbox) ?"
if ($TARGET_FOLDER -eq "") {
    $TARGET_FOLDER = "$PSScriptRoot\vbox"
} else {
    $TARGET_FOLDER = "$TARGET_FOLDER\vbox"
}

# Print configuration
Log-Info "VirtualBox Target Folder: $TARGET_FOLDER"
Log-Verbose "Tools Folder: $TOOLS_FOLDER"
Log-Verbose "Script Root: $PSScriptRoot"

# System Checks
Log-Step "System Checks"

# Check if the target folder exists
Log-Verbose "Checking if VirtualBox folder exists..."
if (-not (Test-Path $TARGET_FOLDER)) {
    Log-Error "The target folder does not exist: $TARGET_FOLDER"
    Log-Info "Please run get_vbox.ps1 first to download and extract VirtualBox."
    exit 1
}
Log-Success "VirtualBox folder found: $TARGET_FOLDER"

# Check if VirtualBox.exe exists
$VBoxExePath = "$TARGET_FOLDER\VirtualBox.exe"
Log-Verbose "Checking for VirtualBox.exe at: $VBoxExePath"
if (-not (Test-Path $VBoxExePath)) {
    Log-Error "VirtualBox.exe not found in: $TARGET_FOLDER"
    Log-Info "The VirtualBox installation may be incomplete."
    exit 1
}
Log-Success "VirtualBox.exe found"

# Get VirtualBox version
Log-Verbose "Attempting to get VirtualBox version..."
$VBoxVersion = $null
try {
    $VBoxInfo = Get-Item $VBoxExePath
    $VBoxVersion = $VBoxInfo.VersionInfo.FileVersion
    if ($VBoxVersion) {
        Log-Info "VirtualBox Version: $VBoxVersion"
    }
} catch {
    Log-Verbose "Could not determine VirtualBox version"
}

# Main
Log-Step "Starting VirtualBox"
Log-Info "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log-Info "Launching VirtualBox from: $VBoxExePath"

try {
    Log-Verbose "Starting VirtualBox process..."
    $process = Start-Process "$VBoxExePath" -PassThru -NoNewWindow
    Log-Success "VirtualBox started (PID: $($process.Id))"
    
    Log-Info "Waiting for VirtualBox to exit..."
    Log-Verbose "You can close this window once VirtualBox is running."
    
    $process.WaitForExit()
    
    $exitCode = $process.ExitCode
    Log-Info "VirtualBox has exited (Exit Code: $exitCode)"
    Log-Info "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    
    if ($exitCode -eq 0) {
        Log-Success "VirtualBox closed normally"
    } else {
        Log-Warning "VirtualBox exited with code: $exitCode"
    }
} catch {
    Log-Error "Failed to start VirtualBox: $($_.Exception.Message)"
    exit 1
}

Log-Success "Script completed"