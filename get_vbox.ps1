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
Write-Host "VBox-PowerPortable v$VERSION - by $AUTHOR`n`n" -ForegroundColor Green

# Constants
$VBOX_WEB_URL = "https://www.virtualbox.org/wiki/Downloads"

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
if ($VERBOSE) {
    Write-Host "[i] Virtualbox Target Folder: $TARGET_FOLDER" -ForegroundColor Green
}

# Functions
function Make-Temp-Dir {
    $TEMP_DIR = [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
    if ($VERBOSE) {
        Write-Host "[i] Creating temporary directory: $TEMP_DIR" -ForegroundColor Green
    }
    New-Item -ItemType Directory -Path $TEMP_DIR -Force | Out-Null
    return $TEMP_DIR
} # Make-Temp-Dir function

function Get-VBox-URL {
    $web_response = Invoke-WebRequest -Uri $VBOX_WEB_URL
    $VBOX_URL = $web_response.Links | Where-Object { $_.class -eq "ext-link" } | ForEach-Object { $_.href } | Select-Object -First 1
    if ($VERBOSE) {
        Write-Host "[i] VirtualBox URL: $VBOX_URL" -ForegroundColor Green
    }
    return $VBOX_URL
} # Get-VBox-URL function

function Get-File-From-URL {
    param (
        [string]$URL,
        [string]$TARGET_FILE
    )
    # check if the file already exists
    if (Test-Path $TARGET_FILE) {
        if ($VERBOSE) {
            Write-Host "[i] File already exists: $TARGET_FILE" -ForegroundColor Green
        }
        throw "File already exists"
        return
    }
    if ($VERBOSE) {
        Write-Host "[i] Getting file from URL: $URL" -ForegroundColor Green
        Write-Host "[i] Writing file to: $TARGET_FILE" -ForegroundColor Green
    }
    Invoke-WebRequest -Uri $URL -OutFile $TARGET_FILE
    if ($VERBOSE) {
        Write-Host "[INFO] File writing completed successfully." -ForegroundColor Green
    }
    return $true
} # Get-File-From-URL function

function Get-Msi-From-Exe {
    param (
        [string]$EXE_FILE,
        [string]$TARGET_FOLDER,
        [string]$TARGET_FILENAME
    )
    # check if the file exists
    if (-not (Test-Path $EXE_FILE)) {
        if ($VERBOSE) {
            Write-Host "[ERROR] File not found: $EXE_FILE" -ForegroundColor Green
        }
        throw "File not found"
        return
    }
    # check if the folder exists
    if (-not (Test-Path $TARGET_FOLDER)) {
        if ($VERBOSE) {
            Write-Host "[ERROR] Folder not found: $TARGET_FOLDER" -ForegroundColor Green
        }
        throw "Folder not found"
        return
    }
    # check if the file already exists
    if (Test-Path "$TARGET_FOLDER\$TARGET_FILENAME") {
        if ($VERBOSE) {
            Write-Host "[ERROR] File already exists: $TARGET_FOLDER\$TARGET_FILENAME" -ForegroundColor Green
        }
        throw "File already exists"
        return
    }

    # extract the MSI file
    $FUNCTION_TEMP_DIR = Make-Temp-Dir
    Start-Process cmd -ArgumentList "/min /c set __COMPAT_LAYER=RUNASINVOKER && start $EXE_FILE --silent -extract -path $FUNCTION_TEMP_DIR" -Wait -NoNewWindow -PassThru | Out-Null
    Move-Item $FUNCTION_TEMP_DIR\VirtualBox-*.msi -Destination "$TARGET_FOLDER\$TARGET_FILENAME" -Force -ErrorAction SilentlyContinue | Out-Null
    Remove-Item -Path $FUNCTION_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
}

function Extract-Msi {
    param (
        [string]$MSI_FILE,
        [string]$TARGET_FOLDER
    )
    # check if the file exists
    if (-not (Test-Path $MSI_FILE)) {
        if ($VERBOSE) {
            Write-Host "[ERROR] File not found: $MSI_FILE" -ForegroundColor Green
        }
        throw "File not found"
        return
    }
    # check if the folder exists
    if (-not (Test-Path $TARGET_FOLDER)) {
        if ($VERBOSE) {
            Write-Host "[ERROR] Folder not found: $TARGET_FOLDER" -ForegroundColor Green
        }
        throw "Folder not found"
        return
    }

    # Extract the MSI file
    $FUNCTION_TEMP_DIR = Make-Temp-Dir
    Start-Process msiexec -ArgumentList "/a $MSI_FILE /qn TARGETDIR=$FUNCTION_TEMP_DIR" -Wait -NoNewWindow -PassThru | Out-Null
    Move-Item $FUNCTION_TEMP_DIR\PFiles\Oracle\VirtualBox\* $TARGET_FOLDER -Force -ErrorAction SilentlyContinue | Out-Null
    Remove-Item -Path $FUNCTION_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    if ($VERBOSE) {
        Write-Host "[i] VirtualBox extracted successfully." -ForegroundColor Green
    }
}

# System Checks
if (-not (Test-Path "C:\Windows\System32\msiexec.exe")) {
    Write-Warning "[ERROR] msiexec.exe not found. Please make sure you have Windows Installer installed."
    exit 1
}

if (Test-Path "$TARGET_FOLDER") {
    Write-Warning "[ERROR] Target folder already exists: $TARGET_FOLDER"
    exit 1
}

# Main
$TEMP_DIR = Make-Temp-Dir
try {
    # Get the latest version of VirtualBox
    $VBOX_URL = Get-VBox-URL
    Write-Host "[i] Getting VirtualBox from: $VBOX_URL" -ForegroundColor Blue
    $VBOX_FILE = Get-File-From-URL -URL $VBOX_URL -TARGET_FILE "$TEMP_DIR\vbox.exe"

    # Extract the EXE file
    Get-Msi-From-Exe -EXE_FILE "$TEMP_DIR\vbox.exe" -TARGET_FOLDER $TEMP_DIR -TARGET_FILENAME "virtualbox.msi"
    if ($VERBOSE) {
        Write-Host "[i] VirtualBox MSI file extracted successfully." -ForegroundColor Green
    }

    # Check if the target folder exists
    if (-not (Test-Path $TARGET_FOLDER)) {
        if ($VERBOSE) {
            Write-Host "[i] Creating target folder: $TARGET_FOLDER" -ForegroundColor Green
        }
        New-Item -ItemType Directory -Path $TARGET_FOLDER -Force | Out-Null
    }

    # Extract the MSI file
    if ($VERBOSE) {
        Write-Host "[i] Extracting VirtualBox MSI file to: $TARGET_FOLDER" -ForegroundColor Green
    }
    Extract-Msi -MSI_FILE "$TEMP_DIR\virtualbox.msi" -TARGET_FOLDER $TARGET_FOLDER

    # Success
    Write-Host "[i] Virtualbox is now located at $TARGET_FOLDER" -ForegroundColor Green

} catch {
    Write-Warning "[ERROR] Failed to get VirtualBox."
    Write-Warning $_.Exception.Message
    exit 1
} finally {
    # cleanup
    if ($VERBOSE) {
        Write-Host "[i] Cleaning up temporary files." -ForegroundColor Green
    }
    Remove-Item -Path $TEMP_DIR -Recurse -Force
}