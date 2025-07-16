# Set Execution Policy
Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Ensure C:\apps\ directory exists
$appsPath = "C:\apps"
if (-not (Test-Path -Path $appsPath)) {
    try {
        New-Item -ItemType Directory -Path $appsPath | Out-Null
        Write-Host "Created directory: $appsPath" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create directory $appsPath. $_"
    }
} else {
    Write-Host "Directory already exists: $appsPath" -ForegroundColor Yellow
}

# Function to prompt user if they want to proceed
function Confirm-Action {
    param (
        [string]$message
    )
    $confirmation = Read-Host "$message Do you want to proceed? (Y/N)"
    if ($confirmation -eq "Y" -or $confirmation -eq "y") {
        return $true
    } else {
        Write-Host "Operation canceled." -ForegroundColor Yellow
        return $false
    }
}

# Function to change PC name based on serial number
function Change-PCName {
    Write-Host "Changing the PC name based on the serial number..." -ForegroundColor Green
    $serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber
    $newPCName = "RPI-" + $serialNumber
    Write-Host "The new PC name will be: $newPCName" -ForegroundColor Yellow
    if (Confirm-Action "This will rename the computer and require a restart.") {
        Rename-Computer -NewName $newPCName -Force -Restart
    }
}

# Function to join the domain
function Join-Domain {
    # Prompt user for domain details
    $domainName = Read-Host "Enter the domain name (e.g., rootprojects.local)"
    $ouPath = Read-Host "Enter the Organizational Unit (OU) path (e.g., OU=Computers,DC=rootprojects,DC=local). Leave blank for default location"
    $credential = Get-Credential -Message "Enter credentials with permission to join the domain"

    # Construct the Add-Computer command
    $command = "Add-Computer -DomainName '$domainName' -Credential \$credential -Force -Restart"

    # If an OU path is provided, append it to the command
    if ($ouPath -ne "") {
        $command += " -OUPath '$ouPath'"
    }

    # Display confirmation and execute the command
    Write-Host "Joining the computer to the domain $domainName..." -ForegroundColor Green
    if (Confirm-Action "This will join the computer to the domain and require a restart.") {
        Invoke-Expression $command
    }
}

# Function to repair Windows using DISM
function Repair-Windows {
    Write-Host "Running DISM command to restore health..." -ForegroundColor Green
    Write-Host "This process can take 15-30 minutes depending on your system and may require a restart." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        dism /online /cleanup-image /restorehealth
        Write-Host "Windows repair completed." -ForegroundColor Green
    }
}

# Function to repair system files using SFC
function Repair-SystemFiles {
    Write-Host "Running System File Checker (SFC)..." -ForegroundColor Green
    Write-Host "This scan can take up to 20 minutes. No restart is required unless issues are found." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        sfc /scannow
        Write-Host "System File Checker completed." -ForegroundColor Green
    }
}

# Function to repair disk using CHKDSK
function Repair-Disk {
    Write-Host "Running Check Disk (CHKDSK)..." -ForegroundColor Green
    Write-Host "This operation could take a few hours and will restart the computer. Ensure all work is saved." -ForegroundColor Red
    if (Confirm-Action "This operation may take several hours and will restart the computer.") {
        chkdsk C: /F /R /X
        Write-Host "Check Disk completed." -ForegroundColor Green
    }
}

# Function to run Windows Update Troubleshooter
function Run-WindowsUpdateTroubleshooter {
    Write-Host "Running Windows Update Troubleshooter..." -ForegroundColor Green
    Write-Host "This operation might take 5-10 minutes. No restart is required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        Start-Process -FilePath "msdt.exe" -ArgumentList "/id WindowsUpdateDiagnostic" -Wait
        Write-Host "Windows Update Troubleshooter completed." -ForegroundColor Green
    }
}

# Function to check and repair DISM
function Check-And-Repair-DISM {
    Write-Host "Running DISM Check and Repair..." -ForegroundColor Green
    Write-Host "This process might take 30-60 minutes depending on system issues. A restart may be required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time and may require a restart.") {
        DISM /Online /Cleanup-Image /CheckHealth
        DISM /Online /Cleanup-Image /ScanHealth
        DISM /Online /Cleanup-Image /RestoreHealth
        Write-Host "DISM Check and Repair completed." -ForegroundColor Green
    }
}

# Function to reset network
function Reset-Network {
    Write-Host "Resetting Network Adapters..." -ForegroundColor Green
    Write-Host "This will cause a temporary network outage and may require a restart." -ForegroundColor Red
    if (Confirm-Action "This operation will temporarily disrupt network connectivity.") {
        netsh winsock reset
        netsh int ip reset
        ipconfig /release
        ipconfig /renew
        ipconfig /flushdns
        Write-Host "Network reset completed. A restart may be required." -ForegroundColor Green
    }
}

# Function to run Windows Memory Diagnostic
function Run-MemoryDiagnostic {
    Write-Host "Running Windows Memory Diagnostic..." -ForegroundColor Green
    Write-Host "This test will restart your computer and may take 15-30 minutes. Ensure all work is saved." -ForegroundColor Red
    if (Confirm-Action "This operation will restart your computer.") {
        Start-Process -FilePath "mdsched.exe" -ArgumentList "/f" -Verb RunAs
        Write-Host "Windows Memory Diagnostic scheduled. The computer will restart to run the test." -ForegroundColor Green
    }
}

# Function to run Startup Repair
function Run-StartupRepair {
    Write-Host "Running Startup Repair..." -ForegroundColor Green
    Write-Host "This process will restart your computer and attempt to fix startup issues. It may take up to an hour." -ForegroundColor Red
    if (Confirm-Action "This operation will restart your computer.") {
        Start-Process -FilePath "reagentc.exe" -ArgumentList "/boottore" -Verb RunAs -Wait
        shutdown /r /t 0
    }
}

# Function to run Windows Defender Full Scan
function Run-WindowsDefenderScan {
    Write-Host "Running Windows Defender Full Scan..." -ForegroundColor Green
    Write-Host "This scan can take several hours depending on the size of your drive. No restart is required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take several hours.") {
        Start-MpScan -ScanType FullScan
        Write-Host "Windows Defender Full Scan completed." -ForegroundColor Green
    }
}

# Function to reset Windows Update components
function Reset-WindowsUpdateComponents {
    Write-Host "Resetting Windows Update components..." -ForegroundColor Green
    Write-Host "This operation may take 10-20 minutes. No restart is required, but Windows Update services will be temporarily unavailable." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        net stop wuauserv
        net stop cryptsvc
        net stop bits
        net stop msiserver
        Rename-Item -Path "C:\Windows\SoftwareDistribution" -NewName "SoftwareDistribution.old" -Force -ErrorAction SilentlyContinue
        Rename-Item -Path "C:\Windows\System32\catroot2" -NewName "Catroot2.old" -Force -ErrorAction SilentlyContinue
        net start wuauserv
        net start cryptsvc
        net start bits
        net start msiserver
        Write-Host "Windows Update components reset completed." -ForegroundColor Green
    }
}

# Function to start Microsoft Teams
function Start-Teams {
    Write-Host "Trying to find and launch Microsoft Teams..." -ForegroundColor Cyan

    # Define possible locations where Teams could be installed
    $possiblePaths = @(
        "$env:LOCALAPPDATA\Programs\Teams\current\Teams.exe",
        "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Local\Microsoft\Teams\current\Teams.exe",
        "$env:PROGRAMFILES\Teams\Teams.exe",
        "$env:PROGRAMFILES(X86)\Teams\Teams.exe",
        "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
        "$env:LOCALAPPDATA\Microsoft\Teams\Teams.exe"
    )

    $teamsPath = $null

    # Search for the Teams executable in possible locations
    foreach ($path in $possiblePaths) {
        try {
            $fullPath = [System.IO.Path]::GetFullPath($path)  # Resolve any relative paths
            if (Test-Path -Path $fullPath) {
                $teamsPath = $fullPath
                Write-Host "Teams executable found at: $teamsPath" -ForegroundColor Green
                break
            } else {
                Write-Host "Teams executable not found at: $fullPath" -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Failed to check path: $path. $_"
        }
    }

    if ($teamsPath) {
        Start-Process -FilePath $teamsPath
    } else {
        Write-Warning "Microsoft Teams executable not found."
    }
}

# Function to clear Teams cache
function Clear-TeamsCache {
    Write-Host "Do you want to delete the Teams Cache (Y/N)?" -ForegroundColor Cyan
    $clearCache = Read-Host "Enter Y to delete the cache or N to cancel"

    if ($clearCache.ToUpper() -eq "Y") {
        Write-Host "Closing Teams" -ForegroundColor Cyan

        try {
            if (Get-Process -ProcessName ms-teams -ErrorAction SilentlyContinue) { 
                Stop-Process -Name ms-teams -Force
                Start-Sleep -Seconds 3
                Write-Host "Teams successfully closed" -ForegroundColor Green
            } else {
                Write-Host "Teams is already closed" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to close Teams. $_"
        }

        Write-Host "Clearing Teams cache" -ForegroundColor Cyan

        $cachePath = "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Local\Microsoft\Teams"

        try {
            # Verify path exists
            if (Test-Path -Path $cachePath) {
                Remove-Item -Path $cachePath -Recurse -Force -Confirm:$false
                Write-Host "Teams cache removed" -ForegroundColor Green
            } else {
                Write-Warning "Teams cache path not found: $cachePath"
            }
        } catch {
            Write-Warning "Failed to remove Teams cache. $_"
        }

        Write-Host "Cleanup complete... Trying to launch Teams" -ForegroundColor Green
        Start-Teams
    } else {
        Write-Host "Cache deletion canceled." -ForegroundColor Yellow
    }
}

# Function to display the Windows Repairs submenu
function Show-WindowsMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "      Windows Repairs       " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: DISM RestoreHealth"
    Write-Host "2: System File Checker (SFC)"
    Write-Host "3: Check Disk (CHKDSK)"
    Write-Host "4: Windows Update Troubleshooter"
    Write-Host "5: DISM Check and Repair"
    Write-Host "6: Network Reset"
    Write-Host "7: Windows Memory Diagnostic"
    Write-Host "8: Windows Startup Repair"
    Write-Host "9: Windows Defender Full Scan"
    Write-Host "10: Reset Windows Update Components"
    Write-Host "11: List installed Apps"
    Write-Host "12: Network Diagnostics"
    Write-Host "13: Factory Reset Device/Reinstall Windows"
    Write-Host "0: Back to Main Menu"
}

# Function to repair Microsoft Office
function Repair-Office {
    Write-Host "Repairing Microsoft Office installation..." -ForegroundColor Green
    $OfficeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $OfficeClickToRunPath) {
        Start-Process -FilePath $OfficeClickToRunPath -ArgumentList "scenario=Repair" -Wait
        Write-Host "Microsoft Office repair completed." -ForegroundColor Green
    } else {
        Write-Host "Microsoft Office Click-to-Run client not found. Please ensure Office is installed." -ForegroundColor Red
    }
}

# Function to check for Microsoft Office updates
function Check-OfficeUpdates {
    Write-Host "Checking for Microsoft Office updates..." -ForegroundColor Green
    $OfficeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $OfficeClickToRunPath) {
        Start-Process -FilePath $OfficeClickToRunPath -ArgumentList "scenario=ApplyUpdates" -Wait
        Write-Host "Office update check completed. Updates have been applied if available." -ForegroundColor Green
    } else {
        Write-Host "Microsoft Office Click-to-Run client not found. Please ensure Office is installed." -ForegroundColor Red
    }
}

# Function to factory reset the device
function Factory-Reset {
    Write-Host "This will reset the system to factory settings." -ForegroundColor Red
    Write-Host "WARNING: All personal files, apps, and settings will be removed." -ForegroundColor Red
    if (Confirm-Action "Do you want to proceed with the factory reset?") {
        try {
            Write-Host "Initiating factory reset..." -ForegroundColor Cyan
            Start-Process -FilePath "systemreset.exe" -ArgumentList "-factoryreset" -Verb RunAs
        } catch {
            Write-Warning "Failed to initiate factory reset. $_"
        }
    } else {
        Write-Host "Factory reset canceled." -ForegroundColor Yellow
    }
}

# Function to update Windows
function Update-Windows {
    Write-Host "Checking for Windows updates..." -ForegroundColor Cyan

    # Check if PSWindowsUpdate module is installed
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Host "PSWindowsUpdate module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force
            Write-Host "PSWindowsUpdate module installed successfully." -ForegroundColor Green
        } catch {
            Write-Warning "Failed to install PSWindowsUpdate module. $_"
            return
        }
    }

    # Import the module
    Import-Module PSWindowsUpdate

    # Perform the update
    try {
        Install-WindowsUpdate -AcceptAll -AutoReboot
        Write-Host "Windows updates installed successfully." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to install Windows updates. $_"
    }
}

# Function to clean up disk space
function Clean-DiskSpace {
    Write-Host "Cleaning up disk space..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait
        Write-Host "Disk cleanup completed." -ForegroundColor Green
    } catch {
        Write-Warning "Disk cleanup failed. $_"
    }
}

# Function to check system health
function Check-SystemHealth {
    Write-Host "Checking system health..." -ForegroundColor Cyan
    try {
        Get-EventLog -LogName System -Newest 10 | Format-Table -AutoSize
        Write-Host "System health check complete." -ForegroundColor Green
    } catch {
        Write-Warning "System health check failed. $_"
    }
}

# Function to backup important files
function Backup-Files {
    $source = Read-Host "Enter the path to the directory you want to back up"
    $destination = Read-Host "Enter the backup destination path"
    Write-Host "Backing up files from $source to $destination..." -ForegroundColor Cyan
    try {
        Copy-Item -Path $source -Destination $destination -Recurse -Force
        Write-Host "Backup completed successfully." -ForegroundColor Green
    } catch {
        Write-Warning "Backup failed. $_"
    }
}

# Function to list installed applications
function List-InstalledApps {
    Write-Host "Listing installed applications..." -ForegroundColor Cyan
    try {
        Get-WmiObject -Class Win32_Product | Select-Object Name, Version | Format-Table -AutoSize
        Write-Host "List of installed applications displayed. This may take some time" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to list installed applications. $_"
    }
}

# Function to perform network diagnostics
function Network-Diagnostics {
    Write-Host "Performing network diagnostics..." -ForegroundColor Cyan

    try {
        # Ping external server
        Write-Host "Pinging external server (google.com)..." -ForegroundColor Cyan
        Test-Connection -ComputerName "google.com" -Count 4 | Format-Table -AutoSize

        # Ping local servers
        $localServers = @("10.60.70.11", "192.168.20.186")
        foreach ($server in $localServers) {
            Write-Host "Pinging local server $server..." -ForegroundColor Cyan
            Test-Connection -ComputerName $server -Count 4 | Format-Table -AutoSize
        }

        Write-Host "Network diagnostics completed." -ForegroundColor Green
    } catch {
        Write-Warning "Network diagnostics failed. $_"
    }
}

# Function to check disk usage
function Check-DiskUsage {
    Write-Host "Checking disk usage..." -ForegroundColor Cyan
    try {
        Get-PSDrive -PSProvider FileSystem | Select-Object Name, 
            @{Name="Used(GB)";Expression={[math]::Round(($_.Used)/1GB,2)}}, 
            @{Name="Free(GB)";Expression={[math]::Round($_.Free/1GB,2)}}, 
            @{Name="Used%";Expression={[math]::Round(($_.Used / ($_.Used + $_.Free)) * 100,2)}} | 
            Format-Table -AutoSize
        Write-Host "Disk usage check completed." -ForegroundColor Green
    } catch {
        Write-Warning "Disk usage check failed. $_"
    }
}

# Function to get system information
function Get-SystemInfo {
    Write-Host "Fetching system information..." -ForegroundColor Cyan
    try {
        Get-ComputerInfo | Select-Object CsName, WindowsVersion, WindowsBuildLabEx, SystemType, TotalPhysicalMemory | Format-Table -AutoSize
        Write-Host "System information displayed." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to fetch system information. $_"
    }
}

# Function to map printer
function Map-Printer {
    param (
        [string]$PrinterIP
    )
    $PrinterName = "\\$PrinterIP"
    Write-Host "Mapping printer at $PrinterIP..." -ForegroundColor Green
    Add-Printer -ConnectionName $PrinterName
    Write-Host "Printer mapped successfully." -ForegroundColor Green
}

# Function to open New PC Files folder
function Open-NewPCFiles {
    Write-Host "Opening New PC Files folder..." -ForegroundColor Green
    Invoke-Expression "explorer.exe '\\RPI-AUS-FS01.rootprojects.local\RPI Admin Archive\Software\RP Files'"
}

# Function to download and open Ninite installer
function Download-And-Open-Ninite {
    Write-Host "Downloading Ninite installer..." -ForegroundColor Green
    $niniteUrl = "https://ninite.com/.net4.8-.net4.8.1-7zip-chrome-vlc-zoom/ninite.exe"
    $outputPath = "$appsPath\ninite.exe"
    try {
        Invoke-WebRequest -Uri $niniteUrl -OutFile $outputPath
        Write-Host "Downloaded Ninite. Opening installer..." -ForegroundColor Green
        Start-Process -FilePath $outputPath
    } catch {
        Write-Warning "Failed to download or open Ninite installer. $_"
    }
}

# Function to download and install MS Teams
function Download-MS-Teams {
    Write-Host "Downloading MS Teams installer and packages..." -ForegroundColor Green
    $teamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $teamsMsixUrl = "https://go.microsoft.com/fwlink/?linkid=2196106"
    $bootstrapperPath = "$appsPath\Teams_bootstrapper.exe"
    $msixPath = "$appsPath\teams.msix"

    try {
        # Download Teams Bootstrapper
        Write-Host "Downloading Teams Bootstrapper..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $teamsBootstrapperUrl -OutFile $bootstrapperPath
        Write-Host "Downloaded Teams Bootstrapper to $bootstrapperPath" -ForegroundColor Green

        # Download Teams MSIX Package
        Write-Host "Downloading Teams MSIX Package..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $teamsMsixUrl -OutFile $msixPath
        Write-Host "Downloaded Teams MSIX Package to $msixPath" -ForegroundColor Green

        # Wait for 5 seconds
        Write-Host "Waiting for 5 seconds before initiating installation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5

        # Run the bootstrapper with specified command
        Write-Host "Running Teams Bootstrapper..." -ForegroundColor Green
        Start-Process -FilePath $bootstrapperPath -ArgumentList "-p -o `"$msixPath`"" -Wait
        Write-Host "Microsoft Teams installation initiated." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to download or install Microsoft Teams. $_"
    }
}

function Download-Agent {
    Write-Host "Downloading Agent..." -ForegroundColor Green
    $AgentUrl = "https://setup.auplatform.connectwise.com/windows/BareboneAgent/32/Main-RP_Infrastructure_Pty_Ltd_Windows_OS_ITSPlatform_TKNe0edb98f-608d-481f-99a3-8bb6465a4f61/MSI/setup"
    $AgentPath = "$appsPath\Agent.msi"

    try {
        # Download FF Agent MSIX Package
        Write-Host "Downloading First Focus Agent..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $AgentUrl -OutFile $AgentPath
        Write-Host "Downloaded Agent MSI Package to $AgentPath" -ForegroundColor Green

        Write-Host "Running First Focus Agent Installer..." -ForegroundColor Green
        Start-Process -FilePath $AgentPath -Wait
        Write-Host "First Focus Agent installation initiated." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to download or install First Focus. $_"
    }
}


# Function to install Adobe Acrobat Reader 32-bit using winget
function Install-AdobeReader {
    Write-Host "Installing Adobe Acrobat Reader 32-bit using winget..." -ForegroundColor Green
    Write-Host "This will download and install Adobe Acrobat Reader. Ensure you have an active internet connection." -ForegroundColor Yellow
    if (Confirm-Action "Do you want to install Adobe Acrobat Reader 32-bit?") {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            try {
                winget install -e --id Adobe.Acrobat.Reader.32-bit -h
                Write-Host "Adobe Acrobat Reader installation initiated." -ForegroundColor Green
            } catch {
                Write-Warning "Failed to install Adobe Acrobat Reader. $_"
            }
        } else {
            Write-Warning "winget is not installed or not found in PATH. Please install winget and try again."
        }
    } else {
        Write-Host "Adobe Acrobat Reader installation canceled." -ForegroundColor Yellow
    }
}

# Function to remove HP bloatware
function Remove-HPBloatware {
    Write-Host "Removing HP bloatware..." -ForegroundColor Green
    if (Confirm-Action "Do you want to proceed with removing HP bloatware?") {

        # List of built-in apps to remove
        $UninstallPackages = @(
            "AD2F1837.HPJumpStarts"
            "AD2F1837.HPPCHardwareDiagnosticsWindows"
            "AD2F1837.HPPowerManager"
            "AD2F1837.HPPrivacySettings"
            "AD2F1837.HPSupportAssistant"
            "AD2F1837.HPSureShieldAI"
            "AD2F1837.HPSystemInformation"
            "AD2F1837.HPQuickDrop"
            "AD2F1837.HPWorkWell"
            "AD2F1837.myHP"
            "AD2F1837.HPDesktopSupportUtilities"
            "AD2F1837.HPQuickTouch"
            "AD2F1837.HPEasyClean"
            "AD2F1837.HPSystemInformation"
        )

        # List of programs to uninstall
        $UninstallPrograms = @(
            "HP Client Security Manager"
            "HP Connection Optimizer"
            "HP Documentation"
            "HP MAC Address Manager"
            "HP Notifications"
            "HP Security Update Service"
            "HP System Default Settings"
            "HP Sure Click"
            "HP Sure Click Security Browser"
            "HP Sure Run"
            "HP Sure Recover"
            "HP Sure Sense"
            "HP Sure Sense Installer"
            "HP Wolf Security"
            "HP Wolf Security Application Support for Sure Sense"
            "HP Wolf Security Application Support for Windows"
        )

        $HPidentifier = "AD2F1837"

        $InstalledPackages = Get-AppxPackage -AllUsers |
            Where-Object { ($UninstallPackages -contains $_.Name) -or ($_.Name -match "^$HPidentifier") }

        $ProvisionedPackages = Get-AppxProvisionedPackage -Online |
            Where-Object { ($UninstallPackages -contains $_.DisplayName) -or ($_.DisplayName -match "^$HPidentifier") }

        $InstalledPrograms = Get-Package | Where-Object { $UninstallPrograms -contains $_.Name }

        # Remove appx provisioned packages
        foreach ($ProvPackage in $ProvisionedPackages) {
            Write-Host "Attempting to remove provisioned package: [$($ProvPackage.DisplayName)]..."
            try {
                Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -ErrorAction Stop | Out-Null
                Write-Host "Successfully removed provisioned package: [$($ProvPackage.DisplayName)]"
            } catch {
                Write-Warning "Failed to remove provisioned package: [$($ProvPackage.DisplayName)]"
            }
        }

        # Remove appx packages
        foreach ($AppxPackage in $InstalledPackages) {
            Write-Host "Attempting to remove Appx package: [$($AppxPackage.Name)]..."
            try {
                Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction Stop | Out-Null
                Write-Host "Successfully removed Appx package: [$($AppxPackage.Name)]"
            } catch {
                Write-Warning "Failed to remove Appx package: [$($AppxPackage.Name)]"
            }
        }

        # Remove installed programs
        foreach ($Program in $InstalledPrograms) {
            Write-Host "Attempting to uninstall: [$($Program.Name)]..."
            try {
                $Program | Uninstall-Package -AllVersions -Force -ErrorAction Stop | Out-Null
                Write-Host "Successfully uninstalled: [$($Program.Name)]"
            } catch {
                Write-Warning "Failed to uninstall: [$($Program.Name)]"
            }
        }

        # Fallback attempt 1 to remove HP Wolf Security using msiexec
        try {
            Start-Process -FilePath "msiexec.exe" -ArgumentList '/x "{0E2E04B0-9EDD-11EB-B38C-10604B96B11E}" /qn /norestart' -Wait
            Write-Host "Fallback to MSI uninstall for HP Wolf Security initiated" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to uninstall HP Wolf Security using MSI - Error message: $($_.Exception.Message)"
        }

        # Fallback attempt 2 to remove HP Wolf Security using msiexec
        try {
            Start-Process -FilePath "msiexec.exe" -ArgumentList '/x "{4DA839F0-72CF-11EC-B247-3863BB3CB5A8}" /qn /norestart' -Wait
            Write-Host "Fallback to MSI uninstall for HP Wolf Security 2 initiated" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to uninstall HP Wolf Security 2 using MSI - Error message: $($_.Exception.Message)"
        }

        Write-Host "HP bloatware removal process completed." -ForegroundColor Green
    } else {
        Write-Host "HP bloatware removal canceled." -ForegroundColor Yellow
    }
}

# Function to install all printers via VBS script
function Install-AllPrinters {
    Write-Host "Installing all printers via VBS script..." -ForegroundColor Green

    # List of possible UNC paths for the VBS script
    $vbsPaths = @(
        "\\RPI-AUS-FS01.rootprojects.local\RPI Admin Archive\Software\RP Files\Printers.vbs"
    )

    $vbsFound = $false

    foreach ($path in $vbsPaths) {
        if (Test-Path $path) {
            Write-Host "Found Printers.vbs at $path" -ForegroundColor Green
            try {
                Start-Process -FilePath "wscript.exe" -ArgumentList "`"$path`"" -Wait
                Write-Host "Printers installation initiated." -ForegroundColor Green
                $vbsFound = $true
                break
            } catch {
                Write-Warning "Failed to run Printers.vbs from $path. $_"
            }
        } else {
            Write-Host "Printers.vbs not found at $path" -ForegroundColor Yellow
        }
    }

    if (-not $vbsFound) {
        Write-Host "Printers.vbs script not found in any of the specified locations." -ForegroundColor Red
    }
}

# Function to display the New PC Setup submenu
function Show-NewPCSetupMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "       New PC Setup         " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: Open New PC Files Folder"
    Write-Host "2: Download and Open Ninite"
    Write-Host "3: Download MS Teams"
    Write-Host "4: Change PC Name"
    Write-Host "5: Join RPI Domain"
    Write-Host "6: Update Windows"
    Write-Host "7: Install Adobe Acrobat Reader 32-bit"
    Write-Host "8: Remove HP Bloatware"
    Write-Host "9: Install First Focus Agent"
    Write-Host "0: Back to Main Menu"
}

# Function to display the Office Repairs submenu
function Show-OfficeMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "       Office Repairs       " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: Repair Microsoft Office"
    Write-Host "2: Check for Microsoft Office Updates"
    Write-Host "0: Back to Main Menu"
}

# Function to display the User Tasks submenu
function Show-UserTasksMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "         User Tasks         " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: Clean Temp Files"
    Write-Host "2: Printer Mapping"
    Write-Host "3: Clear Teams Cache"
    Write-Host "0: Back to Main Menu"
}

# Function to display the Printer Mapping submenu
function Show-PrinterMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "      Printer Mapping       " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: Map Sydney Printer"
    Write-Host "2: Map Melbourne Printer"
    Write-Host "3: Map Melbourne Airport Printer"
    Write-Host "4: Map Townsville Printer"
    Write-Host "5: Map Brisbane Printer"
    Write-Host "6: Map Mackay Printer"
    Write-Host "7: Install All Printers"  # New Option Added
    Write-Host "0: Back to User Tasks Menu"
}

# Function to install temporary files (Clean-TempFiles)
function Clean-TempFiles {
    Write-Host "Cleaning temporary files..." -ForegroundColor Cyan
    $tempPaths = @(
        "$env:TEMP\*",
        "C:\Windows\Temp\*"
    )
    foreach ($path in $tempPaths) {
        try {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Cleared: $path" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to clear: $path. $_"
        }
    }
    Write-Host "Temporary files cleanup completed." -ForegroundColor Green
}


# Function to display the main menu
function Show-MainMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "       RPI Repair Menu      " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: Windows Repairs"
    Write-Host "2: Office Repairs"
    Write-Host "3: User Tasks"
    Write-Host "4: New PC Setup"
    Write-Host "0: Exit"
    Write-Host "===========================" -ForegroundColor DarkYellow
    Write-Host " Jump Functions are enabled " -ForegroundColor DarkYellow
    Write-Host "===========================" -ForegroundColor DarkYellow
}

# Main script loop
do {
    Show-MainMenu
    $mainChoice = Read-Host "Enter your choice (e.g., 1, 1.2, 3.3, etc.)"
    switch ($mainChoice) {
        "1" {
            do {
                Show-WindowsMenu
                $windowsChoice = Read-Host "Enter your choice (0-13)"
                switch ($windowsChoice) {
                    "1" { Repair-Windows }
                    "2" { Repair-SystemFiles }
                    "3" { Repair-Disk }
                    "4" { Run-WindowsUpdateTroubleshooter }
                    "5" { Check-And-Repair-DISM }
                    "6" { Reset-Network }
                    "7" { Run-MemoryDiagnostic }
                    "8" { Run-StartupRepair }
                    "9" { Run-WindowsDefenderScan }
                    "10" { Reset-WindowsUpdateComponents }
                    "11" { List-InstalledApps }
                    "12" { Network-Diagnostics }
                    "13" { Factory-Reset }
                    "0" { break }
                    default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                }
                if ($windowsChoice -ne "0") {
                    Pause
                }
            } while ($windowsChoice -ne "0")
        }
        "2" {
            do {
                Show-OfficeMenu
                $officeChoice = Read-Host "Enter your choice (0-2)"
                switch ($officeChoice) {
                    "1" { Repair-Office }
                    "2" { Check-OfficeUpdates }
                    "0" { break }
                    default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                }
                if ($officeChoice -ne "0") {
                    Pause
                }
            } while ($officeChoice -ne "0")
        }
        "3" {
            do {
                Show-UserTasksMenu
                $userTasksChoice = Read-Host "Enter your choice (0-3)"
                switch ($userTasksChoice) {
                    "1" { Clean-TempFiles }
                    "2" {
                        do {
                            Show-PrinterMenu
                            $printerChoice = Read-Host "Enter your choice (0-7)"
                            switch ($printerChoice) {
                                "1" { Map-Printer -PrinterIP "192.168.23.10" }  # Sydney Printer
                                "2" { Map-Printer -PrinterIP "192.168.33.63" }  # Melbourne Printer
                                "3" { Map-Printer -PrinterIP "192.168.43.250" }  # Melbourne Airport Printer
                                "4" { Map-Printer -PrinterIP "192.168.100.240" }  # Townsville Printer
                                "5" { Map-Printer -PrinterIP "192.168.20.242" }  # Brisbane Printer
                                "6" { Map-Printer -PrinterIP "192.168.90.240" }  # Mackay Printer
                                "7" { Install-AllPrinters }  # Handle new option
                                "0" { break }
                                default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                            }
                            if ($printerChoice -ne "0") {
                                Pause
                            }
                        } while ($printerChoice -ne "0")
                    }
                    "3" { Clear-TeamsCache }
                    "0" { break }
                    default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                }
                if ($userTasksChoice -ne "0") {
                    Pause
                }
            } while ($userTasksChoice -ne "0")
        }
        "4" {
            do {
                Show-NewPCSetupMenu
                $newPCSetupChoice = Read-Host "Enter your choice (0-8)"
                switch ($newPCSetupChoice) {
                    "1" { Open-NewPCFiles }
                    "2" { Download-And-Open-Ninite }
                    "3" { Download-MS-Teams }
                    "4" { Change-PCName }
                    "5" { Join-Domain }
                    "6" { Update-Windows }
                    "7" { Install-AdobeReader }
                    "8" { Remove-HPBloatware }  # Handle new option
                    "9" { Download-Agent }
                    "0" { break }
                    default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                }
                if ($newPCSetupChoice -ne "0") {
                    Pause
                }
            } while ($newPCSetupChoice -ne "0")
        }
        "0" { 
            Write-Host "Exiting..." -ForegroundColor Yellow 
        }
        default { 
            Write-Host "Invalid choice, please try again." -ForegroundColor Red 
        }
    }
    if ($mainChoice -notmatch "^(0|[1-4](\.[1-9]+)?)$") {
        Pause
    }
} while ($mainChoice -ne "0")
