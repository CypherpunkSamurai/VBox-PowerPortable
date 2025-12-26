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
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
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
if ($VERBOSE) {
    Write-Host "[i] Virtualbox Target Folder: $TARGET_FOLDER" -ForegroundColor Green
}

# Functions
function Run-Command {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Command,
        [Parameter(Mandatory=$true)]
        [string]$Arguments
    )

    if ($VERBOSE) {
        Write-Host "[i] Running command:`n  $Command $Arguments" -ForegroundColor Green
    }

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
    $p.Start() | Out-Null
    
    # create a hash table to return both the output and the exit code
    $result = [pscustomobject]@{
        Output = $p.StandardOutput.ReadToEnd()
        Error = $p.StandardError.ReadToEnd()
        ExitCode = $p.ExitCode
    }
    $p.WaitForExit()
    
    # return the object
    return $result
} # Run-Command function

# Check if a Service exists
function Check-SC-ServiceExists {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    # check if the service exists
    $result = Run-Command "sc" "query $ServiceName"

    if ($result.ExitCode -ne 0) {
        return $false
    }

    return $true
}

# Check if a Service is running
function Check-SC-ServiceRunning {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    # check if the service is running
    $result = Run-Command "sc" "query $ServiceName"

    if ($result.Output -match "RUNNING") {
        return $true
    }

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

    # create the service
    $result = Run-Command "sc" "create $ServiceName binPath=`"$ServicePath`" $ServiceArgs"

    if ($result.ExitCode -ne 0) {
        throw "failed to create service. Error: $($result.Error)"
    }

    return $true
}

# Stop a service
function Stop-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    # stop the service
    $result = Run-Command "sc" "stop $ServiceName"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to stop service. Error: $($result.Error)"
    }

    return $true
}

# Start a service
function Start-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    # start the service
    $result = Run-Command "sc" "start $ServiceName"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to start service. Error: $($result.Error)"
    }

    return $true
}

# Remove a service
function Remove-SC-Service {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ServiceName
    )

    # remove the service
    $result = Run-Command "sc" "delete $ServiceName"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to remove service. Error: $($result.Error)"
    }

    return $true
}

# Register a Dll
function Register-Dll {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DllPath
    )

    # register the dll
    $result = Run-Command "regsvr32" "/s `"$DllPath`""

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to register dll. Error: $($result.Error)"
    }

    return $true
}

# Unregister a Dll
function Unregister-Dll {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DllPath
    )

    # unregister the dll
    $result = Run-Command "regsvr32" "/u /s `"$DllPath`""

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to unregister dll. Error: $($result.Error)"
    }

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

    # install the driver
    $result = Run-Command "$DevconPath" "install `"$DriverPath`" $HardwareId"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to install driver. Error: $($result.Error)"
    }

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

    # uninstall the driver
    $result = Run-Command "$DevconPath" "remove $HardwareId"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to uninstall driver. Error: $($result.Error)"
    }

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

    # .\Utils\x64\netcfg.exe -v -l "%cd%\x64\drivers\network\netlwf\VBoxNetLwf.inf" -c s -i "oracle_VBoxNetLwf"
    $result = Run-Command "$netcfgPath" "-v -l `"$NetworkConfigPath`" -c s -i $ConfigID"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to install network config. Error: $($result.Error)"
    }

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

    # .\Utils\x64\netcfg.exe -v -u "oracle_VBoxNetLwf"
    $result = Run-Command "$netcfgPath" "-v -u $ConfigId"

    if ($VERBOSE) {
        Write-Host "$($result.Output)" -ForegroundColor Green
        Write-Host "$($result.Error)" -ForegroundColor Red
    }

    if ($result.ExitCode -ne 0) {
        throw "Failed to uninstall network config. Error: $($result.Error)"
    }

    return $true
}


# System Checks
if (-not (Test-Path "$TARGET_FOLDER")) {
    Write-Warning "[i] VirtualBox is not installed at $TARGET_FOLDER" -ForegroundColor Yellow
    exit
}

# Installer Function
function Install-All-Services {
    param (
        [Parameter(Mandatory=$true)]
        [string]$VBoxFolder
    )

    # Install all services
    # sc create "VBoxSup" binpath="%cd%\x64\drivers\vboxsup\VBoxSup.sys" type=kernel start=system error=normal displayname="VBox SupDrv"
    Create-SC-Service "VBoxSup" "$VBoxFolder\drivers\vboxsup\VBoxSup.sys" "type=kernel start=system error=normal displayname=`"VBox SupDrv`"" -ErrorAction SilentlyContinue
    # sc start "VBoxSup"
    Start-SC-Service "VBoxSup" -ErrorAction SilentlyContinue
    # .\Utils\x64\netcfg.exe -v -l "%cd%\x64\drivers\network\netlwf\VBoxNetLwf.inf" -c s -i "oracle_VBoxNetLwf"
    Install-Network-Config "$TOOLS_FOLDER\x64\netcfg.exe" "$VBoxFolder\drivers\network\netlwf\VBoxNetLwf.inf" "oracle_VBoxNetLwf" -ErrorAction SilentlyContinue
    # sc query "VBoxNetLwf"
    Check-SC-ServiceExists "VBoxNetLwf" -ErrorAction SilentlyContinue
    # .\Utils\x64\devcon.exe install "%cd%\x64\drivers\network\netadp6\VBoxNetAdp6.inf" "sun_VBoxNetAdp"
    Install-Driver "$TOOLS_FOLDER\x64\devcon.exe" "$VBoxFolder\drivers\network\netadp6\VBoxNetAdp6.inf" "sun_VBoxNetAdp" -ErrorAction SilentlyContinue
    # sc query "VBoxNetAdp"
    Check-SC-ServiceExists "VBoxNetAdp" -ErrorAction SilentlyContinue
    # sc create "VBoxUSBMon" binpath="%cd%\x64\drivers\USB\filter\VBoxUSBMon.sys" type=kernel start=system error=normal displayname="VBox USBMon"
    Create-SC-Service "VBoxUSBMon" "$VBoxFolder\drivers\USB\filter\VBoxUSBMon.sys" "type=kernel start=system error=normal displayname=`"VBox USBMon`"" -ErrorAction SilentlyContinue
    # sc start "VBoxUSBMon"
    Start-SC-Service "VBoxUSBMon" -ErrorAction SilentlyContinue
    # .\Utils\x64\devcon.exe install "%cd%\x64\drivers\USB\device\VBoxUSB.inf" "USB\VID_80EE&PID_CAFE"
    Install-Driver "$TOOLS_FOLDER\x64\devcon.exe" "$VBoxFolder\drivers\USB\device\VBoxUSB.inf" "USB\VID_80EE&PID_CAFE" -ErrorAction SilentlyContinue
    # regsvr32.exe /s "%cd%\x64\VBoxProxyStub.dll"
    Register-Dll "$VBoxFolder\VBoxProxyStub.dll" -ErrorAction SilentlyContinue
    # regsvr32.exe /s "%cd%\x64\x86\VBoxProxyStub-x86.dll"
    Register-Dll "$VBoxFolder\x86\VBoxProxyStub-x86.dll" -ErrorAction SilentlyContinue
    # regsvr32.exe /s "%cd%\x64\VBoxC.dll"
    Register-Dll "$VBoxFolder\VBoxC.dll" -ErrorAction SilentlyContinue
    # regsvr32.exe /s "%cd%\x64\x86\VBoxClient-x86.dll"
    Register-Dll "$VBoxFolder\x86\VBoxClient-x86.dll" -ErrorAction SilentlyContinue
    # .\x64\VBoxManage setproperty machinefolder "%cd%\Home\VMs"
    Run-Command "$VBoxFolder\VBoxManage.exe" "setproperty machinefolder `"$PSScriptRoot\Home\VMs`"" -ErrorAction SilentlyContinue
    # ICACLS "%cd%\x64" /setowner "NT Service\TrustedInstaller" /T /L /Q
    Run-Command "ICACLS" "`"$VBoxFolder`" /setowner `"NT Service\TrustedInstaller`" /T /L /Q" -ErrorAction SilentlyContinue
}

# Uninstaller Function
function Uninstall-All-Services {
    param (
        [Parameter(Mandatory=$true)]
        [string]$VBoxFolder
    )

    # regsvr32.exe /s /u .\x64\x86\VBoxClient-x86.dll
    Unregister-Dll "$VBoxFolder\x86\VBoxClient-x86.dll" -ErrorAction SilentlyContinue
    # regsvr32.exe /s /u .\x64\VBoxC.dll
    Unregister-Dll "$VBoxFolder\VBoxC.dll" -ErrorAction SilentlyContinue
    # .\x64\VBoxSDS.exe /UnregService
    Run-Command "$VBoxFolder\VBoxSDS.exe" "/UnregService" -ErrorAction SilentlyContinue
    # regsvr32.exe /s /u .\x64\x86\VBoxProxyStub-x86.dll
    Unregister-Dll "$VBoxFolder\x86\VBoxProxyStub-x86.dll" -ErrorAction SilentlyContinue
    # regsvr32.exe /s /u .\x64\VBoxProxyStub.dll
    Unregister-Dll "$VBoxFolder\VBoxProxyStub.dll" -ErrorAction SilentlyContinue
    # .\Utils\x64\devcon.exe remove "USB\VID_80EE&PID_CAFE"
    Uninstall-Driver "$TOOLS_FOLDER\x64\devcon.exe" "USB\VID_80EE&PID_CAFE" -ErrorAction SilentlyContinue
    # sc delete "VBoxUSB"
    Remove-SC-Service "VBoxUSB" -ErrorAction SilentlyContinue
    # sc stop "VBoxUSBMon"
    Stop-SC-Service "VBoxUSBMon" -ErrorAction SilentlyContinue
    # sc delete "VBoxUSBMon"
    Remove-SC-Service "VBoxUSBMon" -ErrorAction SilentlyContinue
    # sc delete "VBoxNetAdp"
    Remove-SC-Service "VBoxNetAdp" -ErrorAction SilentlyContinue
    # .\Utils\x64\devcon.exe remove "sun_VBoxNetAdp"
    Uninstall-Driver "$TOOLS_FOLDER\x64\devcon.exe" "sun_VBoxNetAdp" -ErrorAction SilentlyContinue
    # sc delete "VBoxNetLwf"
    Remove-SC-Service "VBoxNetLwf" -ErrorAction SilentlyContinue
    # .\Utils\x64\netcfg.exe -v -u "oracle_VBoxNetLwf"
    Uninstall-Network-Config "$TOOLS_FOLDER\x64\netcfg.exe" "oracle_VBoxNetLwf" -ErrorAction SilentlyContinue
    # sc stop "VBoxSup"
    Stop-SC-Service "VBoxSup" -ErrorAction SilentlyContinue
    # sc delete "VBoxSup"
    Remove-SC-Service "VBoxSup" -ErrorAction SilentlyContinue
    # ICACLS "PATH" /setowner "Everyone" /T /L /Q
    Run-Command "ICACLS" "`"$VBoxFolder`" /setowner `"Everyone`" /T /L /Q" -ErrorAction SilentlyContinue
}


# Main
# Menu
Write-Host "[>] Choose an option:"
Write-Host "  1. Install VirtualBox Services"
Write-Host "  2. Uninstall VirtualBox Services"
Write-Host "  q. Exit"
$CHOICE = Read-Host "[>] Enter your choice (1/2):"
if ($CHOICE -eq "1") {
    Install-All-Services $TARGET_FOLDER
    Write-Host "[i] VirtualBox Services installed successfully" -ForegroundColor Blue
} elseif ($CHOICE -eq "2") {
    Uninstall-All-Services $TARGET_FOLDER
    Write-Host "[i] VirtualBox Services uninstalled successfully" -ForegroundColor Blue
} elseif ($CHOICE -eq "q" -or $CHOICE -eq "Q" -or $CHOICE -eq "exit" -or $CHOICE -eq "quit") {
    exit
} else {
    Write-Host "[!] Invalid choice. Exiting..." -ForegroundColor Red
    exit
}