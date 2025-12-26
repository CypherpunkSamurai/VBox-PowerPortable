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
Write-Host "   ~ Setup Script ~`n" -ForegroundColor Green

# Elevate the script
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "[i] Elevating script to Administrator..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"" -Verb RunAs
    exit
}

# Variables
$VERBOSE = $false
# check arguments
if ($args -contains "-v" -or $args -contains "--verbose" -or $args -contains "-verbose" -or $args -contains "--v") {
    Write-Host "[i] Verbose mode enabled" -ForegroundColor Green
    $VERBOSE = $true
}
$TOOLS_FOLDER = "$PSScriptRoot\tools"
$TARGET_FOLDER = Read-Host "[>] Where to setup VirtualBox (default: [CURRENT_FOLDER]\vbox) ?"
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
Log-Info "VirtualBox Target Folder: $TARGET_FOLDER"
Log-Info "Tools Folder: $TOOLS_FOLDER"

# Functions
function Run-Command {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Command,
        [Parameter(Mandatory=$true)]
        [string]$Arguments
    )

    Log-Verbose "Running command: $Command $Arguments"

    # create a process info and assign the properties
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $Command
    $pinfo.Arguments = $Arguments
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.CreateNoWindow = $true
    $pinfo.Arguments = $Arguments

    # create the process object
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    
    Log-Verbose "Starting process..."
    $p.Start() | Out-Null
    
    # create a hash table to return both the output and the exit code
    $result = [pscustomobject]@{
        Output = $p.StandardOutput.ReadToEnd()
        Error = $p.StandardError.ReadToEnd()
        ExitCode = $p.ExitCode
    }
    $p.WaitForExit()
    
    Log-Verbose "Process completed with exit code: $($p.ExitCode)"
    if ($result.Output -and $VERBOSE) {
        Log-Verbose "  Output: $($result.Output.Trim())"
    }
    if ($result.Error -and $VERBOSE) {
        Log-Verbose "  Error: $($result.Error.Trim())"
    }
    
    # return the object
    return $result
} # Run-Command function

# Check if a Service exists
function Check-SC-ServiceExists {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    Log-Verbose "Checking if service exists: $ServiceName"
    $result = Run-Command "sc" "query $ServiceName"

    if ($result.ExitCode -ne 0) {
        Log-Verbose "Service does not exist: $ServiceName"
        return $false
    }

    Log-Verbose "Service exists: $ServiceName"
    return $true
}

# Check if a Service is running
function Check-SC-ServiceRunning {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    Log-Verbose "Checking if service is running: $ServiceName"
    $result = Run-Command "sc" "query $ServiceName"

    if ($result.Output -match "RUNNING") {
        Log-Verbose "Service is running: $ServiceName"
        return $true
    }

    Log-Verbose "Service is not running: $ServiceName"
    return $false
}

# Create a service
function Create-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
        [Parameter(Mandatory=$true)]
        [string]$ServicePath,
        [string]$ServiceArgs=""
    )

    Log-Info "Creating service: $ServiceName"
    Log-Verbose "  Path: $ServicePath"
    Log-Verbose "  Args: $ServiceArgs"
    
    $result = Run-Command "sc" "create $ServiceName binPath=`"$ServicePath`" $ServiceArgs"

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to create service: $ServiceName"
        Log-Verbose "  Error: $($result.Error)"
        throw "failed to create service. Error: $($result.Error)"
    }

    Log-Success "Service created: $ServiceName"
    return $true
}

# Stop a service
function Stop-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    Log-Info "Stopping service: $ServiceName"
    $result = Run-Command "sc" "stop $ServiceName"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to stop service: $ServiceName"
        throw "Failed to stop service. Error: $($result.Error)"
    }

    Log-Success "Service stopped: $ServiceName"
    return $true
}

# Start a service
function Start-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    Log-Info "Starting service: $ServiceName"
    $result = Run-Command "sc" "start $ServiceName"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to start service: $ServiceName"
        throw "Failed to start service. Error: $($result.Error)"
    }

    Log-Success "Service started: $ServiceName"
    return $true
}

# Remove a service
function Remove-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    Log-Info "Removing service: $ServiceName"
    $result = Run-Command "sc" "delete $ServiceName"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to remove service: $ServiceName"
        throw "Failed to remove service. Error: $($result.Error)"
    }

    Log-Success "Service removed: $ServiceName"
    return $true
}

# Register a Dll
function Register-Dll {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DllPath
    )

    Log-Info "Registering DLL: $DllPath"
    
    if (-not (Test-Path $DllPath)) {
        Log-Warning "DLL not found: $DllPath"
        return $false
    }
    
    $result = Run-Command "regsvr32" "/s `"$DllPath`""

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to register DLL: $DllPath"
        throw "Failed to register dll. Error: $($result.Error)"
    }

    Log-Success "DLL registered: $DllPath"
    return $true
}

# Unregister a Dll
function Unregister-Dll {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DllPath
    )

    Log-Info "Unregistering DLL: $DllPath"
    
    if (-not (Test-Path $DllPath)) {
        Log-Warning "DLL not found: $DllPath"
        return $false
    }
    
    $result = Run-Command "regsvr32" "/u /s `"$DllPath`""

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to unregister DLL: $DllPath"
        throw "Failed to unregister dll. Error: $($result.Error)"
    }

    Log-Success "DLL unregistered: $DllPath"
    return $true
}

# Install Driver
function Install-Driver {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DevconPath,
        [Parameter(Mandatory=$true)]
        [string]$DriverPath,
        [Parameter(Mandatory=$true)]
        [string]$HardwareId
    )

    Log-Info "Installing driver: $HardwareId"
    Log-Verbose "  DevCon: $DevconPath"
    Log-Verbose "  Driver INF: $DriverPath"
    
    if (-not (Test-Path $DevconPath)) {
        Log-Error "DevCon not found: $DevconPath"
        throw "DevCon not found"
    }
    
    if (-not (Test-Path $DriverPath)) {
        Log-Error "Driver INF not found: $DriverPath"
        throw "Driver INF not found"
    }
    
    $result = Run-Command "$DevconPath" "install `"$DriverPath`" $HardwareId"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to install driver: $HardwareId"
        throw "Failed to install driver. Error: $($result.Error)"
    }

    Log-Success "Driver installed: $HardwareId"
    return $true
}

# Uninstall Driver
function Uninstall-Driver {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DevconPath,
        [Parameter(Mandatory=$true)]
        [string]$HardwareId
    )

    Log-Info "Uninstalling driver: $HardwareId"
    Log-Verbose "  DevCon: $DevconPath"
    
    if (-not (Test-Path $DevconPath)) {
        Log-Error "DevCon not found: $DevconPath"
        throw "DevCon not found"
    }
    
    $result = Run-Command "$DevconPath" "remove $HardwareId"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to uninstall driver: $HardwareId"
        throw "Failed to uninstall driver. Error: $($result.Error)"
    }

    Log-Success "Driver uninstalled: $HardwareId"
    return $true
}

# Network Config Install
function Install-Network-Config {
    param (
        [Parameter(Mandatory=$true)]
        [string]$netcfgPath,
        [Parameter(Mandatory=$true)]
        [string]$NetworkConfigPath,
        [Parameter(Mandatory=$true)]
        [string]$ConfigId
    )

    Log-Info "Installing network config: $ConfigId"
    Log-Verbose "  NetCfg: $netcfgPath"
    Log-Verbose "  Config INF: $NetworkConfigPath"
    
    if (-not (Test-Path $netcfgPath)) {
        Log-Error "NetCfg not found: $netcfgPath"
        throw "NetCfg not found"
    }
    
    if (-not (Test-Path $NetworkConfigPath)) {
        Log-Error "Network config INF not found: $NetworkConfigPath"
        throw "Network config INF not found"
    }
    
    # .\Utils\x64\netcfg.exe -v -l "%cd%\x64\drivers\network\netlwf\VBoxNetLwf.inf" -c s -i "oracle_VBoxNetLwf"
    $result = Run-Command "$netcfgPath" "-v -l `"$NetworkConfigPath`" -c s -i $ConfigID"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to install network config: $ConfigId"
        throw "Failed to install network config. Error: $($result.Error)"
    }

    Log-Success "Network config installed: $ConfigId"
    return $true
}

# Network Config Uninstall
function Uninstall-Network-Config {
    param (
        [Parameter(Mandatory=$true)]
        [string]$netcfgPath,
        [Parameter(Mandatory=$true)]
        [string]$ConfigId
    )

    Log-Info "Uninstalling network config: $ConfigId"
    Log-Verbose "  NetCfg: $netcfgPath"
    
    if (-not (Test-Path $netcfgPath)) {
        Log-Error "NetCfg not found: $netcfgPath"
        throw "NetCfg not found"
    }
    
    # .\Utils\x64\netcfg.exe -v -u "oracle_VBoxNetLwf"
    $result = Run-Command "$netcfgPath" "-v -u $ConfigId"

    if ($VERBOSE) {
        if ($result.Output) { Log-Verbose "  Output: $($result.Output.Trim())" }
        if ($result.Error) { Log-Verbose "  Error: $($result.Error.Trim())" }
    }

    if ($result.ExitCode -ne 0) {
        Log-Error "Failed to uninstall network config: $ConfigId"
        throw "Failed to uninstall network config. Error: $($result.Error)"
    }

    Log-Success "Network config uninstalled: $ConfigId"
    return $true
}


# System Checks
Log-Step "System Checks"

Log-Verbose "Checking if VirtualBox is installed at target folder..."
if (-not (Test-Path "$TARGET_FOLDER")) {
    Log-Error "VirtualBox is not installed at: $TARGET_FOLDER"
    Log-Info "Please run get_vbox.ps1 first to download and extract VirtualBox."
    exit 1
}
Log-Success "VirtualBox folder found: $TARGET_FOLDER"

Log-Verbose "Checking for tools folder..."
if (-not (Test-Path "$TOOLS_FOLDER")) {
    Log-Warning "Tools folder not found: $TOOLS_FOLDER"
} else {
    Log-Success "Tools folder found: $TOOLS_FOLDER"
}

# Installer Function
function Install-All-Services {
    param (
        [Parameter(Mandatory=$true)]
        [string]$VBoxFolder
    )

    Log-Step "Installing VirtualBox Services"
    Log-Info "VirtualBox folder: $VBoxFolder"
    Log-Info "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    
    $StepNumber = 1
    
    # Install all services
    # sc create "VBoxSup" binpath="%cd%\x64\drivers\vboxsup\VBoxSup.sys" type=kernel start=system error=normal displayname="VBox SupDrv"
    Log-Info "Step $StepNumber`: Creating VBoxSup kernel driver service..."
    $StepNumber++
    Create-SC-Service "VBoxSup" "$VBoxFolder\drivers\vboxsup\VBoxSup.sys" "type=kernel start=system error=normal displayname=`"VBox SupDrv`"" -ErrorAction SilentlyContinue
    
    # sc start "VBoxSup"
    Log-Info "Step $StepNumber`: Starting VBoxSup service..."
    $StepNumber++
    Start-SC-Service "VBoxSup" -ErrorAction SilentlyContinue
    
    # .\Utils\x64\netcfg.exe -v -l "%cd%\x64\drivers\network\netlwf\VBoxNetLwf.inf" -c s -i "oracle_VBoxNetLwf"
    Log-Info "Step $StepNumber`: Installing VBoxNetLwf network config..."
    $StepNumber++
    Install-Network-Config "$TOOLS_FOLDER\x64\netcfg.exe" "$VBoxFolder\drivers\network\netlwf\VBoxNetLwf.inf" "oracle_VBoxNetLwf" -ErrorAction SilentlyContinue
    
    # sc query "VBoxNetLwf"
    Log-Info "Step $StepNumber`: Verifying VBoxNetLwf service..."
    $StepNumber++
    $lwfExists = Check-SC-ServiceExists "VBoxNetLwf" -ErrorAction SilentlyContinue
    if ($lwfExists) { Log-Success "VBoxNetLwf service exists" } else { Log-Warning "VBoxNetLwf service not found" }
    
    # .\Utils\x64\devcon.exe install "%cd%\x64\drivers\network\netadp6\VBoxNetAdp6.inf" "sun_VBoxNetAdp"
    Log-Info "Step $StepNumber`: Installing VBoxNetAdp6 driver..."
    $StepNumber++
    Install-Driver "$TOOLS_FOLDER\x64\devcon.exe" "$VBoxFolder\drivers\network\netadp6\VBoxNetAdp6.inf" "sun_VBoxNetAdp" -ErrorAction SilentlyContinue
    
    # sc query "VBoxNetAdp"
    Log-Info "Step $StepNumber`: Verifying VBoxNetAdp service..."
    $StepNumber++
    $adpExists = Check-SC-ServiceExists "VBoxNetAdp" -ErrorAction SilentlyContinue
    if ($adpExists) { Log-Success "VBoxNetAdp service exists" } else { Log-Warning "VBoxNetAdp service not found" }
    
    # sc create "VBoxUSBMon" binpath="%cd%\x64\drivers\USB\filter\VBoxUSBMon.sys" type=kernel start=system error=normal displayname="VBox USBMon"
    Log-Info "Step $StepNumber`: Creating VBoxUSBMon kernel driver service..."
    $StepNumber++
    Create-SC-Service "VBoxUSBMon" "$VBoxFolder\drivers\USB\filter\VBoxUSBMon.sys" "type=kernel start=system error=normal displayname=`"VBox USBMon`"" -ErrorAction SilentlyContinue
    
    # sc start "VBoxUSBMon"
    Log-Info "Step $StepNumber`: Starting VBoxUSBMon service..."
    $StepNumber++
    Start-SC-Service "VBoxUSBMon" -ErrorAction SilentlyContinue
    
    # .\Utils\x64\devcon.exe install "%cd%\x64\drivers\USB\device\VBoxUSB.inf" "USB\VID_80EE&PID_CAFE"
    Log-Info "Step $StepNumber`: Installing VBoxUSB driver..."
    $StepNumber++
    Install-Driver "$TOOLS_FOLDER\x64\devcon.exe" "$VBoxFolder\drivers\USB\device\VBoxUSB.inf" "USB\VID_80EE&PID_CAFE" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s "%cd%\x64\VBoxProxyStub.dll"
    Log-Info "Step $StepNumber`: Registering VBoxProxyStub.dll..."
    $StepNumber++
    Register-Dll "$VBoxFolder\VBoxProxyStub.dll" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s "%cd%\x64\x86\VBoxProxyStub-x86.dll"
    Log-Info "Step $StepNumber`: Registering VBoxProxyStub-x86.dll..."
    $StepNumber++
    Register-Dll "$VBoxFolder\x86\VBoxProxyStub-x86.dll" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s "%cd%\x64\VBoxC.dll"
    Log-Info "Step $StepNumber`: Registering VBoxC.dll..."
    $StepNumber++
    Register-Dll "$VBoxFolder\VBoxC.dll" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s "%cd%\x64\x86\VBoxClient-x86.dll"
    Log-Info "Step $StepNumber`: Registering VBoxClient-x86.dll..."
    $StepNumber++
    Register-Dll "$VBoxFolder\x86\VBoxClient-x86.dll" -ErrorAction SilentlyContinue
    
    # .\x64\VBoxManage setproperty machinefolder "%cd%\Home\VMs"
    Log-Info "Step $StepNumber`: Setting VirtualBox machine folder..."
    $StepNumber++
    $vmFolder = "$PSScriptRoot\Home\VMs"
    Log-Verbose "Setting machine folder to: $vmFolder"
    Run-Command "$VBoxFolder\VBoxManage.exe" "setproperty machinefolder `"$vmFolder`"" -ErrorAction SilentlyContinue
    
    # ICACLS "%cd%\x64" /setowner "NT Service\TrustedInstaller" /T /L /Q
    Log-Info "Step $StepNumber`: Setting folder ownership to TrustedInstaller..."
    $StepNumber++
    Run-Command "ICACLS" "`"$VBoxFolder`" /setowner `"NT Service\TrustedInstaller`" /T /L /Q" -ErrorAction SilentlyContinue
    
    Log-Step "Installation Completed"
    Log-Info "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Log-Success "All services installed successfully!"
}

# Uninstaller Function
function Uninstall-All-Services {
    param (
        [Parameter(Mandatory=$true)]
        [string]$VBoxFolder
    )

    Log-Step "Uninstalling VirtualBox Services"
    Log-Info "VirtualBox folder: $VBoxFolder"
    Log-Info "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    
    $StepNumber = 1

    # regsvr32.exe /s /u .\x64\x86\VBoxClient-x86.dll
    Log-Info "Step $StepNumber`: Unregistering VBoxClient-x86.dll..."
    $StepNumber++
    Unregister-Dll "$VBoxFolder\x86\VBoxClient-x86.dll" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s /u .\x64\VBoxC.dll
    Log-Info "Step $StepNumber`: Unregistering VBoxC.dll..."
    $StepNumber++
    Unregister-Dll "$VBoxFolder\VBoxC.dll" -ErrorAction SilentlyContinue
    
    # .\x64\VBoxSDS.exe /UnregService
    Log-Info "Step $StepNumber`: Unregistering VBoxSDS service..."
    $StepNumber++
    Run-Command "$VBoxFolder\VBoxSDS.exe" "/UnregService" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s /u .\x64\x86\VBoxProxyStub-x86.dll
    Log-Info "Step $StepNumber`: Unregistering VBoxProxyStub-x86.dll..."
    $StepNumber++
    Unregister-Dll "$VBoxFolder\x86\VBoxProxyStub-x86.dll" -ErrorAction SilentlyContinue
    
    # regsvr32.exe /s /u .\x64\VBoxProxyStub.dll
    Log-Info "Step $StepNumber`: Unregistering VBoxProxyStub.dll..."
    $StepNumber++
    Unregister-Dll "$VBoxFolder\VBoxProxyStub.dll" -ErrorAction SilentlyContinue
    
    # .\Utils\x64\devcon.exe remove "USB\VID_80EE&PID_CAFE"
    Log-Info "Step $StepNumber`: Removing VBoxUSB driver..."
    $StepNumber++
    Uninstall-Driver "$TOOLS_FOLDER\x64\devcon.exe" "USB\VID_80EE&PID_CAFE" -ErrorAction SilentlyContinue
    
    # sc delete "VBoxUSB"
    Log-Info "Step $StepNumber`: Removing VBoxUSB service..."
    $StepNumber++
    Remove-SC-Service "VBoxUSB" -ErrorAction SilentlyContinue
    
    # sc stop "VBoxUSBMon"
    Log-Info "Step $StepNumber`: Stopping VBoxUSBMon service..."
    $StepNumber++
    Stop-SC-Service "VBoxUSBMon" -ErrorAction SilentlyContinue
    
    # sc delete "VBoxUSBMon"
    Log-Info "Step $StepNumber`: Removing VBoxUSBMon service..."
    $StepNumber++
    Remove-SC-Service "VBoxUSBMon" -ErrorAction SilentlyContinue
    
    # sc delete "VBoxNetAdp"
    Log-Info "Step $StepNumber`: Removing VBoxNetAdp service..."
    $StepNumber++
    Remove-SC-Service "VBoxNetAdp" -ErrorAction SilentlyContinue
    
    # .\Utils\x64\devcon.exe remove "sun_VBoxNetAdp"
    Log-Info "Step $StepNumber`: Removing VBoxNetAdp driver..."
    $StepNumber++
    Uninstall-Driver "$TOOLS_FOLDER\x64\devcon.exe" "sun_VBoxNetAdp" -ErrorAction SilentlyContinue
    
    # sc delete "VBoxNetLwf"
    Log-Info "Step $StepNumber`: Removing VBoxNetLwf service..."
    $StepNumber++
    Remove-SC-Service "VBoxNetLwf" -ErrorAction SilentlyContinue
    
    # .\Utils\x64\netcfg.exe -v -u "oracle_VBoxNetLwf"
    Log-Info "Step $StepNumber`: Uninstalling VBoxNetLwf network config..."
    $StepNumber++
    Uninstall-Network-Config "$TOOLS_FOLDER\x64\netcfg.exe" "oracle_VBoxNetLwf" -ErrorAction SilentlyContinue
    
    # sc stop "VBoxSup"
    Log-Info "Step $StepNumber`: Stopping VBoxSup service..."
    $StepNumber++
    Stop-SC-Service "VBoxSup" -ErrorAction SilentlyContinue
    
    # sc delete "VBoxSup"
    Log-Info "Step $StepNumber`: Removing VBoxSup service..."
    $StepNumber++
    Remove-SC-Service "VBoxSup" -ErrorAction SilentlyContinue
    
    # ICACLS "PATH" /setowner "Everyone" /T /L /Q
    Log-Info "Step $StepNumber`: Resetting folder ownership to Everyone..."
    $StepNumber++
    Run-Command "ICACLS" "`"$VBoxFolder`" /setowner `"Everyone`" /T /L /Q" -ErrorAction SilentlyContinue
    
    Log-Step "Uninstallation Completed"
    Log-Info "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Log-Success "All services uninstalled successfully!"
}


# Main
# Menu
Log-Step "Main Menu"
Write-Host "[>] Choose an option:"
Write-Host "  1. Install VirtualBox Services"
Write-Host "  2. Uninstall VirtualBox Services"
Write-Host "  q. Exit"
$CHOICE = Read-Host "[>] Enter your choice (1/2)"

if ($CHOICE -eq "1") {
    Install-All-Services $TARGET_FOLDER
} elseif ($CHOICE -eq "2") {
    Uninstall-All-Services $TARGET_FOLDER
} elseif ($CHOICE -eq "q" -or $CHOICE -eq "Q" -or $CHOICE -eq "exit" -or $CHOICE -eq "quit") {
    Log-Info "Exiting..."
    exit
} else {
    Log-Error "Invalid choice. Exiting..."
    exit 1
}