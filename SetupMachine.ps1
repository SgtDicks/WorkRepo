Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

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

# Define the functions for each task
function Change-PCName {
    Write-Host "Changing the PC name based on the serial number..." -ForegroundColor Green
    $serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber
    $newPCName = "RPI-" + $serialNumber
    Write-Host "The new PC name will be: $newPCName" -ForegroundColor Yellow
    if (Confirm-Action "This will rename the computer and require a restart.") {
        Rename-Computer -NewName $newPCName -Force -Restart
    }
}

function Join-Domain {
    Write-Host "Joining the computer to the domain rootprojects.local..." -ForegroundColor Green
    if (Confirm-Action "This will join the computer to the domain and require a restart.") {
        Add-Computer -DomainName "rootprojects.local" -Credential (Get-Credential) -Restart
    }
}

function Repair-Windows {
    Write-Host "Running DISM command to restore health..." -ForegroundColor Green
    Write-Host "This process can take 15-30 minutes depending on your system and may require a restart." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        dism /online /cleanup-image /restorehealth
        Write-Host "Windows repair completed." -ForegroundColor Green
    }
}

function Repair-SystemFiles {
    Write-Host "Running System File Checker (SFC)..." -ForegroundColor Green
    Write-Host "This scan can take up to 20 minutes. No restart is required unless issues are found." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        sfc /scannow
        Write-Host "System File Checker completed." -ForegroundColor Green
    }
}

function Repair-Disk {
    Write-Host "Running Check Disk (CHKDSK)..." -ForegroundColor Green
    Write-Host "This operation could take a few hours and will restart the computer. Ensure all work is saved." -ForegroundColor Red
    if (Confirm-Action "This operation may take several hours and will restart the computer.") {
        chkdsk C: /F /R /X
        Write-Host "Check Disk completed." -ForegroundColor Green
    }
}

function Run-WindowsUpdateTroubleshooter {
    Write-Host "Running Windows Update Troubleshooter..." -ForegroundColor Green
    Write-Host "This operation might take 5-10 minutes. No restart is required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        Start-Process -FilePath "msdt.exe" -ArgumentList "/id WindowsUpdateDiagnostic" -Wait
        Write-Host "Windows Update Troubleshooter completed." -ForegroundColor Green
    }
}

function CheckAndRepair-DISM {
    Write-Host "Running DISM Check and Repair..." -ForegroundColor Green
    Write-Host "This process might take 30-60 minutes depending on system issues. A restart may be required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time and may require a restart.") {
        DISM /Online /Cleanup-Image /CheckHealth
        DISM /Online /Cleanup-Image /ScanHealth
        DISM /Online /Cleanup-Image /RestoreHealth
        Write-Host "DISM Check and Repair completed." -ForegroundColor Green
    }
}

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

function Run-MemoryDiagnostic {
    Write-Host "Running Windows Memory Diagnostic..." -ForegroundColor Green
    Write-Host "This test will restart your computer and may take 15-30 minutes. Ensure all work is saved." -ForegroundColor Red
    if (Confirm-Action "This operation will restart your computer.") {
        Start-Process -FilePath "mdsched.exe" -ArgumentList "/f" -Verb RunAs
        Write-Host "Windows Memory Diagnostic scheduled. The computer will restart to run the test." -ForegroundColor Green
    }
}

function Run-StartupRepair {
    Write-Host "Running Startup Repair..." -ForegroundColor Green
    Write-Host "This process will restart your computer and attempt to fix startup issues. It may take up to an hour." -ForegroundColor Red
    if (Confirm-Action "This operation will restart your computer.") {
        Start-Process -FilePath "reagentc.exe" -ArgumentList "/boottore" -Verb RunAs -Wait
        shutdown /r /t 0
    }
}

function Run-WindowsDefenderScan {
    Write-Host "Running Windows Defender Full Scan..." -ForegroundColor Green
    Write-Host "This scan can take several hours depending on the size of your drive. No restart is required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take several hours.") {
        Start-MpScan -ScanType FullScan
        Write-Host "Windows Defender Full Scan completed." -ForegroundColor Green
    }
}

function Reset-WindowsUpdateComponents {
    Write-Host "Resetting Windows Update components..." -ForegroundColor Green
    Write-Host "This operation may take 10-20 minutes. No restart is required, but Windows Update services will be temporarily unavailable." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        net stop wuauserv
        net stop cryptsvc
        net stop bits
        net stop msiserver
        Ren C:\Windows\SoftwareDistribution SoftwareDistribution.old
        Ren C:\Windows\System32\catroot2 Catroot2.old
        net start wuauserv
        net start cryptsvc
        net start bits
        net start msiserver
        Write-Host "Windows Update components reset completed." -ForegroundColor Green
    }
}
function Start-Teams {
    Write-Host "Trying to find and launch Microsoft Teams..." -ForegroundColor Cyan

    # Define possible locations where Teams could be installed
    $possiblePaths = @(
        "$env:LOCALAPPDATA\Programs\Teams\current\Teams.exe",
        "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Local\Microsoft\Teams\current\Teams.exe",
        "$env:PROGRAMFILES\Teams\Teams.exe",
        "$env:PROGRAMFILES(X86)\Teams\Teams.exe"
    )

    $teamsPath = $null

    # Search for the Teams executable in possible locations
    foreach ($path in $possiblePaths) {
        if (Test-Path -Path $path) {
            $teamsPath = $path
            break
        }
    }

    if ($teamsPath) {
        Write-Host "Teams executable found at: $teamsPath" -ForegroundColor Green
        Start-Process -FilePath $teamsPath
    } else {
        Write-Warning "Microsoft Teams executable not found."
    }
}

function Clear-TeamsCache {
    Write-Host "Do you want to delete the Teams Cache (Y/N)?" -ForegroundColor Cyan
    $clearCache = Read-Host "Do you want to delete the Teams Cache (Y/N)?"

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
            Write-Warning $_
        }

        Write-Host "Clearing Teams cache" -ForegroundColor Cyan

        try {
            Remove-Item -Path "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Local\Microsoft\Teams" -Recurse -Force -Confirm:$false
            Write-Host "Teams cache removed" -ForegroundColor Green
        } catch {
            Write-Warning $_
        }

        Write-Host "Cleanup complete... Trying to launch Teams" -ForegroundColor Green
        Start-Teams
    }
}

# Function to display the Windows Repairs submenu
function Show-WindowsMenu {
    Clear-Host
    Write-Host "Windows Repairs Menu" -ForegroundColor Cyan
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
    Write-Host "12: Network Test"
    Write-Host "0: Back to Main Menu"
}

# Define the functions for other tasks
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
# Update Windows
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

# Clean Up Disk Space
function Clean-DiskSpace {
    Write-Host "Cleaning up disk space..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait
        Write-Host "Disk cleanup completed." -ForegroundColor Green
    } catch {
        Write-Warning "Disk cleanup failed. $_"
    }
}

# Check System Health
function Check-SystemHealth {
    Write-Host "Checking system health..." -ForegroundColor Cyan
    try {
        Get-EventLog -LogName System -Newest 10 | Format-Table -AutoSize
        Write-Host "System health check complete." -ForegroundColor Green
    } catch {
        Write-Warning "System health check failed. $_"
    }
}

# Backup Important Files
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

# List Installed Applications
function List-InstalledApps {
    Write-Host "Listing installed applications..." -ForegroundColor Cyan
    try {
        Get-WmiObject -Class Win32_Product | Select-Object Name, Version | Format-Table -AutoSize
        Write-Host "List of installed applications displayed. This may take some time" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to list installed applications. $_"
    }
}

# Network Diagnostics
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

# Check Disk Usage
function Check-DiskUsage {
    Write-Host "Checking disk usage..." -ForegroundColor Cyan
    try {
        Get-PSDrive -PSProvider FileSystem | Select-Object Name, @{Name="Used(GB)";Expression={[math]::Round($_.Used/1GB,2)}}, @{Name="Free(GB)";Expression={[math]::Round($_.Free/1GB,2)}}, @{Name="Used%";Expression={[math]::Round($_.Used/$_.Used*100,2)}} | Format-Table -AutoSize
        Write-Host "Disk usage check completed." -ForegroundColor Green
    } catch {
        Write-Warning "Disk usage check failed. $_"
    }
}

# System Information
function Get-SystemInfo {
    Write-Host "Fetching system information..." -ForegroundColor Cyan
    try {
        Get-ComputerInfo | Select-Object CsName, WindowsVersion, WindowsBuildLabEx, SystemType, TotalPhysicalMemory | Format-Table -AutoSize
        Write-Host "System information displayed." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to fetch system information. $_"
    }
}
function Map-Printer {
    param (
        [string]$PrinterIP
    )
    $PrinterName = "\\$PrinterIP"
    Write-Host "Mapping printer at $PrinterIP..." -ForegroundColor Green
    Add-Printer -ConnectionName $PrinterName
    Write-Host "Printer mapped successfully." -ForegroundColor Green
}

function Open-NewPCFiles {
    Write-Host "Opening New PC Files folder..." -ForegroundColor Green
    Invoke-Expression "explorer.exe '\\server-syd\Scans\do not delete this folder\new pc files'"
}

function Download-And-Open-Ninite {
    Write-Host "Downloading Ninite installer..." -ForegroundColor Green
    $niniteUrl = "https://ninite.com/.net4.8-.net4.8.1-7zip-chrome-vlc-zoom/ninite.exe"
    $outputPath = "C:\apps\ninite.exe"
    Invoke-WebRequest -Uri $niniteUrl -OutFile $outputPath
    Write-Host "Downloaded Ninite. Opening installer..." -ForegroundColor Green
    Start-Process -FilePath $outputPath
}

function Download-MS-Teams {
    Write-Host "Downloading MS Teams installer..." -ForegroundColor Green
    $teamsUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $outputPath = "C:\apps\Teams_bootstrapper.exe"
    Invoke-WebRequest -Uri $teamsUrl -OutFile $outputPath
    Write-Host "Downloaded MS Teams installer. Running installer as Admin..." -ForegroundColor Green
    Start-Process -FilePath $outputPath -ArgumentList "-p" -Verb RunAs
}

# Function to display the main menu
function Show-MainMenu {
    Clear-Host
    Write-Host "RPI Repair Menu" -ForegroundColor Cyan
    Write-Host "1: Windows Repairs"
    Write-Host "2: Office Repairs"
    Write-Host "3: User Tasks"
    Write-Host "4: New PC Setup"
    Write-Host "0: Exit"
}

# Function to display the Office Repairs submenu
function Show-OfficeMenu {
    Clear-Host
    Write-Host "Office Repairs Menu" -ForegroundColor Cyan
    Write-Host "1: Repair Microsoft Office"
    Write-Host "2: Check for Microsoft Office Updates"
    Write-Host "0: Back to Main Menu"
}

# Function to display the User Tasks submenu
function Show-UserTasksMenu {
    Clear-Host
    Write-Host "User Tasks Menu" -ForegroundColor Cyan
    Write-Host "1: Clean Temp Files"
    Write-Host "2: Printer Mapping"
    Write-Host "3: Clear Teams Cache"
    Write-Host "0: Back to Main Menu"
}

# Function to display the Printer Mapping submenu under User Tasks
function Show-PrinterMenu {
    Clear-Host
    Write-Host "Printer Mapping Menu" -ForegroundColor Cyan
    Write-Host "1: Map Sydney Printer"
    Write-Host "2: Map Melbourne Printer"
    Write-Host "3: Map Melbourne Airport Printer"
    Write-Host "4: Map Townsville Printer"
    Write-Host "5: Map Brisbane Printer"
    Write-Host "6: Map Mackay Printer"
    Write-Host "0: Back to User Tasks Menu"
}

# Function to display the New PC Setup submenu
function Show-NewPCSetupMenu {
    Clear-Host
    Write-Host "New PC Setup Menu" -ForegroundColor Cyan
    Write-Host "1: Open New PC Files Folder"
    Write-Host "2: Download and Open Ninite"
    Write-Host "3: Download MS Teams"
    Write-Host "4: Change PC Name"
    Write-Host "5: Join RPI Domain"
    Write-Host "6: Update Windows"
    Write-Host "0: Back to Main Menu"
}

# Main script loop
do {
    Show-MainMenu
    $mainChoice = Read-Host "Enter your choice (0-4)"
    switch ($mainChoice) {
        "1" {
            do {
                Show-WindowsMenu
                $windowsChoice = Read-Host "Enter your choice (0-10)"
                switch ($windowsChoice) {
                    "1" { Repair-Windows }
                    "2" { Repair-SystemFiles }
                    "3" { Repair-Disk }
                    "4" { Run-WindowsUpdateTroubleshooter }
                    "5" { CheckAndRepair-DISM }
                    "6" { Reset-Network }
                    "7" { Run-MemoryDiagnostic }
                    "8" { Run-StartupRepair }
                    "9" { Run-WindowsDefenderScan }
                    "10" { Reset-WindowsUpdateComponents }
                    "11" { List-InstalledApps }
                    "12" { Network-Diagnostics }
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
                                $printerChoice = Read-Host "Enter your choice (0-6)"
                                switch ($printerChoice) {
                                 "1" { Map-Printer -PrinterIP "192.168.23.10" }  # Sydney Printer
                                 "2" { Map-Printer -PrinterIP "192.168.33.63" }  # Melbourne Printer
                                 "3" { Map-Printer -PrinterIP "192.168.43.250" }  # Melbourne Airport Printer
                                 "4" { Map-Printer -PrinterIP "192.168.100.240" }  # Townsville Printer
                                 "5" { Map-Printer -PrinterIP "192.168.20.242" }  # Brisbane Printer
                                 "6" { Map-Printer -PrinterIP "192.168.90.240" }  # Mackay Printer
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
                $newPCSetupChoice = Read-Host "Enter your choice (0-5)"
                switch ($newPCSetupChoice) {
                    "1" { Open-NewPCFiles }
                    "2" { Download-And-Open-Ninite }
                    "3" { Download-MS-Teams }
                    "4" { Change-PCName }
                    "5" { Join-Domain }
                    "6" { Update-Windows }
                    "0" { break }
                    default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                }
                if ($newPCSetupChoice -ne "0") {
                    Pause
                }
            } while ($newPCSetupChoice -ne "0")
        }
        "0" { Write-Host "Exiting..." -ForegroundColor Yellow }
        default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
    }
    if ($mainChoice -ne "0" -and $mainChoice -ne "1" -and $mainChoice -ne "2" -and $mainChoice -ne "3" -and $mainChoice -ne "4") {
        Pause
    }
} while ($mainChoice -ne "0")
