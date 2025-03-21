# ===============================
# RPI Repair and Maintenance Tool
# ===============================

# Set Execution Policy and ensure C:\apps\ exists
Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# -------------------------------
# Helper and Utility Functions
# -------------------------------

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

# -------------------------------
# PC and Domain Functions
# -------------------------------

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
    $domainName = Read-Host "Enter the domain name (e.g., rootprojects.local)"
    $ouPath = Read-Host "Enter the Organizational Unit (OU) path (e.g., OU=Computers,DC=rootprojects,DC=local). Leave blank for default location"
    $credential = Get-Credential -Message "Enter credentials with permission to join the domain"
    $command = "Add-Computer -DomainName '$domainName' -Credential \$credential -Force -Restart"
    if ($ouPath -ne "") {
        $command += " -OUPath '$ouPath'"
    }
    Write-Host "Joining the computer to the domain $domainName..." -ForegroundColor Green
    if (Confirm-Action "This will join the computer to the domain and require a restart.") {
        Invoke-Expression $command
    }
}

# -------------------------------
# Windows Repair Functions
# -------------------------------

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
        Rename-Item -Path "C:\Windows\SoftwareDistribution" -NewName "SoftwareDistribution.old" -Force -ErrorAction SilentlyContinue
        Rename-Item -Path "C:\Windows\System32\catroot2" -NewName "Catroot2.old" -Force -ErrorAction SilentlyContinue
        net start wuauserv
        net start cryptsvc
        net start bits
        net start msiserver
        Write-Host "Windows Update components reset completed." -ForegroundColor Green
    }
}

# -------------------------------
# Office and Teams Functions
# -------------------------------

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

function Download-MS-Teams {
    Write-Host "Downloading MS Teams installer and packages..." -ForegroundColor Green
    $teamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $teamsMsixUrl = "https://go.microsoft.com/fwlink/?linkid=2196106"
    $bootstrapperPath = "$appsPath\Teams_bootstrapper.exe"
    $msixPath = "$appsPath\teams.msix"
    try {
        Write-Host "Downloading Teams Bootstrapper..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $teamsBootstrapperUrl -OutFile $bootstrapperPath
        Write-Host "Downloaded Teams Bootstrapper to $bootstrapperPath" -ForegroundColor Green
        Write-Host "Downloading Teams MSIX Package..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $teamsMsixUrl -OutFile $msixPath
        Write-Host "Downloaded Teams MSIX Package to $msixPath" -ForegroundColor Green
        Write-Host "Waiting for 5 seconds before initiating installation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        Write-Host "Running Teams Bootstrapper..." -ForegroundColor Green
        Start-Process -FilePath $bootstrapperPath -ArgumentList "-p -o `"$msixPath`"" -Wait
        Write-Host "Microsoft Teams installation initiated." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to download or install Microsoft Teams. $_"
    }
}

function Start-Teams {
    Write-Host "Trying to find and launch Microsoft Teams..." -ForegroundColor Cyan
    $possiblePaths = @(
        "$env:LOCALAPPDATA\Programs\Teams\current\Teams.exe",
        "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Local\Microsoft\Teams\current\Teams.exe",
        "$env:PROGRAMFILES\Teams\Teams.exe",
        "$env:PROGRAMFILES(X86)\Teams\Teams.exe",
        "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
        "$env:LOCALAPPDATA\Microsoft\Teams\Teams.exe"
    )
    $teamsPath = $null
    foreach ($path in $possiblePaths) {
        try {
            $fullPath = [System.IO.Path]::GetFullPath($path)
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

# -------------------------------
# Update, Cleanup, and System Functions
# -------------------------------

function Update-Windows {
    Write-Host "Checking for Windows updates..." -ForegroundColor Cyan
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
    Import-Module PSWindowsUpdate
    try {
        Install-WindowsUpdate -AcceptAll -AutoReboot
        Write-Host "Windows updates installed successfully." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to install Windows updates. $_"
    }
}

function Clean-DiskSpace {
    Write-Host "Cleaning up disk space..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait
        Write-Host "Disk cleanup completed." -ForegroundColor Green
    } catch {
        Write-Warning "Disk cleanup failed. $_"
    }
}

function Check-SystemHealth {
    Write-Host "Checking system health..." -ForegroundColor Cyan
    try {
        Get-EventLog -LogName System -Newest 10 | Format-Table -AutoSize
        Write-Host "System health check complete." -ForegroundColor Green
    } catch {
        Write-Warning "System health check failed. $_"
    }
}

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

function List-InstalledApps {
    Write-Host "Listing installed applications..." -ForegroundColor Cyan
    try {
        Get-WmiObject -Class Win32_Product | Select-Object Name, Version | Format-Table -AutoSize
        Write-Host "List of installed applications displayed. This may take some time." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to list installed applications. $_"
    }
}

function Network-Diagnostics {
    Write-Host "Performing network diagnostics..." -ForegroundColor Cyan
    try {
        Write-Host "Pinging external server (google.com)..." -ForegroundColor Cyan
        Test-Connection -ComputerName "google.com" -Count 4 | Format-Table -AutoSize
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

function Get-SystemInfo {
    Write-Host "Fetching system information..." -ForegroundColor Cyan
    try {
        Get-ComputerInfo | Select-Object CsName, WindowsVersion, WindowsBuildLabEx, SystemType, TotalPhysicalMemory | Format-Table -AutoSize
        Write-Host "System information displayed." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to fetch system information. $_"
    }
}

# -------------------------------
# Printer and New PC Setup Functions
# -------------------------------

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
    $outputPath = "$appsPath\ninite.exe"
    try {
        Invoke-WebRequest -Uri $niniteUrl -OutFile $outputPath
        Write-Host "Downloaded Ninite. Opening installer..." -ForegroundColor Green
        Start-Process -FilePath $outputPath
    } catch {
        Write-Warning "Failed to download or open Ninite installer. $_"
    }
}

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

function Remove-HPBloatware {
    Write-Host "Removing HP bloatware and crapware..." -ForegroundColor Green
    if (Confirm-Action "Do you want to proceed with removing HP bloatware?") {
        $UninstallPackages = @(
            "AD2F1837.HPJumpStarts",
            "AD2F1837.HPPCHardwareDiagnosticsWindows",
            "AD2F1837.HPPowerManager",
            "AD2F1837.HPPrivacySettings",
            "AD2F1837.HPSupportAssistant",
            "AD2F1837.HPSureShieldAI",
            "AD2F1837.HPSystemInformation",
            "AD2F1837.HPQuickDrop",
            "AD2F1837.HPWorkWell",
            "AD2F1837.myHP",
            "AD2F1837.HPDesktopSupportUtilities",
            "AD2F1837.HPQuickTouch",
            "AD2F1837.HPEasyClean",
            "AD2F1837.HPSystemInformation"
        )
        $UninstallPrograms = @(
            "HP Client Security Manager",
            "HP Connection Optimizer",
            "HP Documentation",
            "HP MAC Address Manager",
            "HP Notifications",
            "HP Security Update Service",
            "HP System Default Settings",
            "HP Sure Click",
            "HP Sure Click Security Browser",
            "HP Sure Run",
            "HP Sure Recover",
            "HP Sure Sense",
            "HP Sure Sense Installer",
            "HP Wolf Security",
            "HP Wolf Security Application Support for Sure Sense",
            "HP Wolf Security Application Support for Windows"
        )
        $HPidentifier = "AD2F1837"
        $InstalledPackages = Get-AppxPackage -AllUsers |
            Where-Object { ($UninstallPackages -contains $_.Name) -or ($_.Name -match "^$HPidentifier") }
        $ProvisionedPackages = Get-AppxProvisionedPackage -Online |
            Where-Object { ($UninstallPackages -contains $_.DisplayName) -or ($_.DisplayName -match "^$HPidentifier") }
        $InstalledPrograms = Get-Package | Where-Object { $UninstallPrograms -contains $_.Name }
        foreach ($ProvPackage in $ProvisionedPackages) {
            Write-Host "Attempting to remove provisioned package: [$($ProvPackage.DisplayName)]..."
            try {
                Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -ErrorAction Stop | Out-Null
                Write-Host "Successfully removed provisioned package: [$($ProvPackage.DisplayName)]"
            } catch {
                Write-Warning "Failed to remove provisioned package: [$($ProvPackage.DisplayName)]"
            }
        }
        foreach ($AppxPackage in $InstalledPackages) {
            Write-Host "Attempting to remove Appx package: [$($AppxPackage.Name)]..."
            try {
                Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction Stop | Out-Null
                Write-Host "Successfully removed Appx package: [$($AppxPackage.Name)]"
            } catch {
                Write-Warning "Failed to remove Appx package: [$($AppxPackage.Name)]"
            }
        }
        foreach ($Program in $InstalledPrograms) {
            Write-Host "Attempting to uninstall: [$($Program.Name)]..."
            try {
                $Program | Uninstall-Package -AllVersions -Force -ErrorAction Stop | Out-Null
                Write-Host "Successfully uninstalled: [$($Program.Name)]"
            } catch {
                Write-Warning "Failed to uninstall: [$($Program.Name)]"
            }
        }
        try {
            Start-Process -FilePath "msiexec.exe" -ArgumentList '/x "{0E2E04B0-9EDD-11EB-B38C-10604B96B11E}" /qn /norestart' -Wait
            Write-Host "Fallback to MSI uninstall for HP Wolf Security initiated" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to uninstall HP Wolf Security using MSI - Error message: $($_.Exception.Message)"
        }
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

function Install-AllPrinters {
    Write-Host "Installing all printers via VBS script..." -ForegroundColor Green
    $vbsPaths = @(
        "\\server-mel\software\rp files\Printers.vbs",
        "\\server-syd\Scans\do not delete this folder\new pc files\Printers.vbs"
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

# -------------------------------
# GUI Main Menu Functions
# -------------------------------

function Show-MainMenuGUI {
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "RPI Repair Menu"
    $mainForm.Size = New-Object System.Drawing.Size(800,600)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $mainForm.MaximizeBox = $false

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = "RPI Repair and Maintenance Tool"
    $headerLabel.Font = New-Object System.Drawing.Font("Arial",16,[System.Drawing.FontStyle]::Bold)
    $headerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $headerLabel.Size = New-Object System.Drawing.Size(760,40)
    $headerLabel.Location = New-Object System.Drawing.Point(20,20)
    $headerLabel.ForeColor = [System.Drawing.Color]::FromArgb(0,102,204)
    $mainForm.Controls.Add($headerLabel)

    $buttonWidth = 170
    $buttonHeight = 120
    $horizontalSpacing = 20
    $startX = 20
    $startY = 80
    $currentX = $startX
    $currentY = $startY

    # Windows Repairs button
    $windowsButton = New-Object System.Windows.Forms.Button
    $windowsButton.Text = "Windows Repairs"
    $windowsButton.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
    $windowsButton.Location = New-Object System.Drawing.Point($currentX,$currentY)
    $windowsButton.BackColor = [System.Drawing.Color]::LightBlue
    $windowsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $windowsButton.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $windowsButton.Add_Click({ Show-WindowsMenuGUI })
    $mainForm.Controls.Add($windowsButton)

    $currentX += $buttonWidth + $horizontalSpacing

    # Office Repairs button
    $officeButton = New-Object System.Windows.Forms.Button
    $officeButton.Text = "Office Repairs"
    $officeButton.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
    $officeButton.Location = New-Object System.Drawing.Point($currentX,$currentY)
    $officeButton.BackColor = [System.Drawing.Color]::LightGreen
    $officeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $officeButton.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $officeButton.Add_Click({ Show-OfficeMenuGUI })
    $mainForm.Controls.Add($officeButton)

    $currentX += $buttonWidth + $horizontalSpacing

    # User Tasks button
    $userTasksButton = New-Object System.Windows.Forms.Button
    $userTasksButton.Text = "User Tasks"
    $userTasksButton.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
    $userTasksButton.Location = New-Object System.Drawing.Point($currentX,$currentY)
    $userTasksButton.BackColor = [System.Drawing.Color]::LightYellow
    $userTasksButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $userTasksButton.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $userTasksButton.Add_Click({ Show-UserTasksMenuGUI })
    $mainForm.Controls.Add($userTasksButton)

    $currentX += $buttonWidth + $horizontalSpacing

    # New PC Setup button
    $newPCButton = New-Object System.Windows.Forms.Button
    $newPCButton.Text = "New PC Setup"
    $newPCButton.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
    $newPCButton.Location = New-Object System.Drawing.Point($currentX,$currentY)
    $newPCButton.BackColor = [System.Drawing.Color]::LightPink
    $newPCButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $newPCButton.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $newPCButton.Add_Click({ Show-NewPCSetupMenuGUI })
    $mainForm.Controls.Add($newPCButton)

    # Second row buttons
    $currentX = $startX
    $currentY += $buttonHeight + 20

    # Network Tools button
    $networkButton = New-Object System.Windows.Forms.Button
    $networkButton.Text = "Network Tools"
    $networkButton.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
    $networkButton.Location = New-Object System.Drawing.Point($currentX,$currentY)
    $networkButton.BackColor = [System.Drawing.Color]::LightCyan
    $networkButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $networkButton.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $networkButton.Add_Click({ Network-Diagnostics })
    $mainForm.Controls.Add($networkButton)

    $currentX += $buttonWidth + $horizontalSpacing

    # Power Management button
    $powerButton = New-Object System.Windows.Forms.Button
    $powerButton.Text = "Power Management"
    $powerButton.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
    $powerButton.Location = New-Object System.Drawing.Point($currentX,$currentY)
    $powerButton.BackColor = [System.Drawing.Color]::LightSalmon
    $powerButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $powerButton.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
    $powerButton.Add_Click({ Show-PowerManagementMenuGUI })
    $mainForm.Controls.Add($powerButton)

    [void]$mainForm.ShowDialog()
}

function Show-WindowsMenuGUI {
    $windowsForm = New-Object System.Windows.Forms.Form
    $windowsForm.Text = "Windows Repairs"
    $windowsForm.Size = New-Object System.Drawing.Size(600,500)
    $windowsForm.StartPosition = "CenterScreen"
    $windowsForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $windowsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $windowsForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Windows Repair Options"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(560,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $windowsForm.Controls.Add($titleLabel)
    
    $buttonWidth = 170; $buttonHeight = 60; $buttonsPerRow = 3; $horizontalSpacing = 15; $verticalSpacing = 10
    $startX = 20; $startY = 60; $currentX = $startX; $currentY = $startY; $buttonCount = 0
    $windowsOptions = @(
        @{ Text = "DISM RestoreHealth"; Action = { Repair-Windows } },
        @{ Text = "System File Checker"; Action = { Repair-SystemFiles } },
        @{ Text = "Check Disk"; Action = { Repair-Disk } },
        @{ Text = "Windows Update Troubleshooter"; Action = { Run-WindowsUpdateTroubleshooter } },
        @{ Text = "DISM Check and Repair"; Action = { Check-And-Repair-DISM } },
        @{ Text = "Network Reset"; Action = { Reset-Network } },
        @{ Text = "Memory Diagnostic"; Action = { Run-MemoryDiagnostic } },
        @{ Text = "Startup Repair"; Action = { Run-StartupRepair } },
        @{ Text = "Defender Full Scan"; Action = { Run-WindowsDefenderScan } },
        @{ Text = "Reset Update Components"; Action = { Reset-WindowsUpdateComponents } },
        @{ Text = "List Installed Apps"; Action = { List-InstalledApps } },
        @{ Text = "Factory Reset"; Action = { Factory-Reset } }
    )
    
    foreach ($option in $windowsOptions) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $option.Text
        $button.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
        $button.Location = New-Object System.Drawing.Point($currentX,$currentY)
        $button.BackColor = [System.Drawing.Color]::LightBlue
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Tag = $option.Action
        $button.Add_Click({ param($sender, $e) & $sender.Tag })
        $windowsForm.Controls.Add($button)
        $buttonCount++
        if ($buttonCount % $buttonsPerRow -eq 0) {
            $currentX = $startX; $currentY += $buttonHeight + $verticalSpacing
        } else {
            $currentX += $buttonWidth + $horizontalSpacing
        }
    }
    
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back to Main Menu"
    $backButton.Size = New-Object System.Drawing.Size(150,40)
    $backButton.Location = New-Object System.Drawing.Point(225,410)
    $backButton.BackColor = [System.Drawing.Color]::LightGray
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $backButton.Add_Click({ $windowsForm.Close() })
    $windowsForm.Controls.Add($backButton)
    
    [void]$windowsForm.ShowDialog()
}

function Show-OfficeMenuGUI {
    $officeForm = New-Object System.Windows.Forms.Form
    $officeForm.Text = "Office Repairs"
    $officeForm.Size = New-Object System.Drawing.Size(400,250)
    $officeForm.StartPosition = "CenterScreen"
    $officeForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $officeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $officeForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Microsoft Office Repair Options"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(360,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $officeForm.Controls.Add($titleLabel)

    $repairButton = New-Object System.Windows.Forms.Button
    $repairButton.Text = "Repair Microsoft Office"
    $repairButton.Size = New-Object System.Drawing.Size(200,40)
    $repairButton.Location = New-Object System.Drawing.Point(100,70)
    $repairButton.BackColor = [System.Drawing.Color]::LightGreen
    $repairButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $repairButton.Add_Click({ Repair-Office })
    $officeForm.Controls.Add($repairButton)

    $updateButton = New-Object System.Windows.Forms.Button
    $updateButton.Text = "Check for Office Updates"
    $updateButton.Size = New-Object System.Drawing.Size(200,40)
    $updateButton.Location = New-Object System.Drawing.Point(100,120)
    $updateButton.BackColor = [System.Drawing.Color]::LightGreen
    $updateButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $updateButton.Add_Click({ Check-OfficeUpdates })
    $officeForm.Controls.Add($updateButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back to Main Menu"
    $backButton.Size = New-Object System.Drawing.Size(150,30)
    $backButton.Location = New-Object System.Drawing.Point(125,180)
    $backButton.BackColor = [System.Drawing.Color]::LightGray
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $backButton.Add_Click({ $officeForm.Close() })
    $officeForm.Controls.Add($backButton)
    
    [void]$officeForm.ShowDialog()
}

function Show-UserTasksMenuGUI {
    $userTasksForm = New-Object System.Windows.Forms.Form
    $userTasksForm.Text = "User Tasks"
    $userTasksForm.Size = New-Object System.Drawing.Size(400,350)
    $userTasksForm.StartPosition = "CenterScreen"
    $userTasksForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $userTasksForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $userTasksForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "User Tasks and Maintenance"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(360,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $userTasksForm.Controls.Add($titleLabel)

    $cleanTempButton = New-Object System.Windows.Forms.Button
    $cleanTempButton.Text = "Clean Temporary Files"
    $cleanTempButton.Size = New-Object System.Drawing.Size(200,40)
    $cleanTempButton.Location = New-Object System.Drawing.Point(100,70)
    $cleanTempButton.BackColor = [System.Drawing.Color]::LightYellow
    $cleanTempButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cleanTempButton.Add_Click({ Clean-TempFiles })
    $userTasksForm.Controls.Add($cleanTempButton)

    $printerButton = New-Object System.Windows.Forms.Button
    $printerButton.Text = "Printer Mapping"
    $printerButton.Size = New-Object System.Drawing.Size(200,40)
    $printerButton.Location = New-Object System.Drawing.Point(100,120)
    $printerButton.BackColor = [System.Drawing.Color]::LightYellow
    $printerButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $printerButton.Add_Click({ Show-PrinterMenuGUI })
    $userTasksForm.Controls.Add($printerButton)

    $teamsButton = New-Object System.Windows.Forms.Button
    $teamsButton.Text = "Clear Teams Cache"
    $teamsButton.Size = New-Object System.Drawing.Size(200,40)
    $teamsButton.Location = New-Object System.Drawing.Point(100,170)
    $teamsButton.BackColor = [System.Drawing.Color]::LightYellow
    $teamsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $teamsButton.Add_Click({ Clear-TeamsCache })
    $userTasksForm.Controls.Add($teamsButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back to Main Menu"
    $backButton.Size = New-Object System.Drawing.Size(150,30)
    $backButton.Location = New-Object System.Drawing.Point(125,280)
    $backButton.BackColor = [System.Drawing.Color]::LightGray
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $backButton.Add_Click({ $userTasksForm.Close() })
    $userTasksForm.Controls.Add($backButton)

    [void]$userTasksForm.ShowDialog()
}

function Show-PrinterMenuGUI {
    $printerForm = New-Object System.Windows.Forms.Form
    $printerForm.Text = "Printer Mapping"
    $printerForm.Size = New-Object System.Drawing.Size(450,400)
    $printerForm.StartPosition = "CenterScreen"
    $printerForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $printerForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $printerForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Printer Mapping Options"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(410,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $printerForm.Controls.Add($titleLabel)

    $printerOptions = @(
        @{ Text = "Map Sydney Printer"; IP = "192.168.23.10" },
        @{ Text = "Map Melbourne Printer"; IP = "192.168.33.63" },
        @{ Text = "Map Melbourne Airport Printer"; IP = "192.168.43.250" },
        @{ Text = "Map Townsville Printer"; IP = "192.168.100.240" },
        @{ Text = "Map Brisbane Printer"; IP = "192.168.20.242" },
        @{ Text = "Map Mackay Printer"; IP = "192.168.90.240" }
    )
    $buttonY = 70
    foreach ($option in $printerOptions) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $option.Text
        $button.Size = New-Object System.Drawing.Size(220,35)
        $button.Location = New-Object System.Drawing.Point(115,$buttonY)
        $button.BackColor = [System.Drawing.Color]::Azure
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Tag = $option.IP
        $button.Add_Click({ param($sender,$e) Map-Printer -PrinterIP $sender.Tag })
        $printerForm.Controls.Add($button)
        $buttonY += 45
    }

    $allPrintersButton = New-Object System.Windows.Forms.Button
    $allPrintersButton.Text = "Install All Printers"
    $allPrintersButton.Size = New-Object System.Drawing.Size(220,35)
    $allPrintersButton.Location = New-Object System.Drawing.Point(115,$buttonY)
    $allPrintersButton.BackColor = [System.Drawing.Color]::LightGreen
    $allPrintersButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $allPrintersButton.Add_Click({ Install-AllPrinters })
    $printerForm.Controls.Add($allPrintersButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back"
    $backButton.Size = New-Object System.Drawing.Size(100,30)
    $backButton.Location = New-Object System.Drawing.Point(175,340)
    $backButton.BackColor = [System.Drawing.Color]::LightGray
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $backButton.Add_Click({ $printerForm.Close() })
    $printerForm.Controls.Add($backButton)
    
    [void]$printerForm.ShowDialog()
}

function Show-NewPCSetupMenuGUI {
    $setupForm = New-Object System.Windows.Forms.Form
    $setupForm.Text = "New PC Setup"
    $setupForm.Size = New-Object System.Drawing.Size(500,350)
    $setupForm.StartPosition = "CenterScreen"
    $setupForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $setupForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $setupForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "New PC Setup Options"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(460,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $setupForm.Controls.Add($titleLabel)

    $options = @(
        @{ Text = "Open New PC Files Folder"; Action = { Open-NewPCFiles } },
        @{ Text = "Download and Open Ninite"; Action = { Download-And-Open-Ninite } },
        @{ Text = "Download MS Teams"; Action = { Download-MS-Teams } },
        @{ Text = "Change PC Name"; Action = { Change-PCName } },
        @{ Text = "Join RPI Domain"; Action = { Join-Domain } },
        @{ Text = "Update Windows"; Action = { Update-Windows } },
        @{ Text = "Install Adobe Reader 32-bit"; Action = { Install-AdobeReader } },
        @{ Text = "Remove HP Bloatware"; Action = { Remove-HPBloatware } }
    )
    $buttonY = 70
    foreach ($opt in $options) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $opt.Text
        $button.Size = New-Object System.Drawing.Size(220,35)
        $button.Location = New-Object System.Drawing.Point(140,$buttonY)
        $button.BackColor = [System.Drawing.Color]::LightBlue
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Add_Click({ param($sender,$e) & $opt.Action })
        $setupForm.Controls.Add($button)
        $buttonY += 45
    }

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back to Main Menu"
    $backButton.Size = New-Object System.Drawing.Size(150,30)
    $backButton.Location = New-Object System.Drawing.Point(175,300)
    $backButton.BackColor = [System.Drawing.Color]::LightGray
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $backButton.Add_Click({ $setupForm.Close() })
    $setupForm.Controls.Add($backButton)
    
    [void]$setupForm.ShowDialog()
}

function Show-PowerManagementMenuGUI {
    $powerForm = New-Object System.Windows.Forms.Form
    $powerForm.Text = "Power Management"
    $powerForm.Size = New-Object System.Drawing.Size(400,300)
    $powerForm.StartPosition = "CenterScreen"
    $powerForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Power Management Options"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(360,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $powerForm.Controls.Add($titleLabel)
    
    $batteryButton = New-Object System.Windows.Forms.Button
    $batteryButton.Text = "Check Battery Health"
    $batteryButton.Size = New-Object System.Drawing.Size(200,40)
    $batteryButton.Location = New-Object System.Drawing.Point(100,70)
    $batteryButton.BackColor = [System.Drawing.Color]::LightGreen
    $batteryButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $batteryButton.Add_Click({ param($sender,$e) Check-BatteryHealth })
    $powerForm.Controls.Add($batteryButton)
    
    $powerPlanButton = New-Object System.Windows.Forms.Button
    $powerPlanButton.Text = "Manage Power Plans"
    $powerPlanButton.Size = New-Object System.Drawing.Size(200,40)
    $powerPlanButton.Location = New-Object System.Drawing.Point(100,120)
    $powerPlanButton.BackColor = [System.Drawing.Color]::LightGreen
    $powerPlanButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $powerPlanButton.Add_Click({ param($sender,$e) Manage-PowerPlans })
    $powerForm.Controls.Add($powerPlanButton)
    
    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back"
    $backButton.Size = New-Object System.Drawing.Size(100,30)
    $backButton.Location = New-Object System.Drawing.Point(150,200)
    $backButton.BackColor = [System.Drawing.Color]::LightGray
    $backButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $backButton.Add_Click({ $powerForm.Close() })
    $powerForm.Controls.Add($backButton)
    
    [void]$powerForm.ShowDialog()
}

# (Optional: You could also add a Check-BatteryHealth GUI function if needed)

# -------------------------------
# Console Menu Functions
# -------------------------------

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

function Show-OfficeMenu {
    Clear-Host
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "       Office Repairs       " -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "1: Repair Microsoft Office"
    Write-Host "2: Check for Microsoft Office Updates"
    Write-Host "0: Back to Main Menu"
}

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
    Write-Host "7: Install All Printers"
    Write-Host "0: Back to User Tasks Menu"
}

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
    Write-Host "0: Back to Main Menu"
}

# -------------------------------
# Main Script Loop
# -------------------------------

# Detect if GUI mode is available
$guiMode = $true
try { [System.Windows.Forms.Application] | Out-Null } catch { $guiMode = $false }

if ($guiMode) {
    Show-MainMenuGUI
} else {
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
                    if ($windowsChoice -ne "0") { Pause }
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
                    if ($officeChoice -ne "0") { Pause }
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
                                    "1" { Map-Printer -PrinterIP "192.168.23.10" }
                                    "2" { Map-Printer -PrinterIP "192.168.33.63" }
                                    "3" { Map-Printer -PrinterIP "192.168.43.250" }
                                    "4" { Map-Printer -PrinterIP "192.168.100.240" }
                                    "5" { Map-Printer -PrinterIP "192.168.20.242" }
                                    "6" { Map-Printer -PrinterIP "192.168.90.240" }
                                    "7" { Install-AllPrinters }
                                    "0" { break }
                                    default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                                }
                                if ($printerChoice -ne "0") { Pause }
                            } while ($printerChoice -ne "0")
                        }
                        "3" { Clear-TeamsCache }
                        "0" { break }
                        default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                    }
                    if ($userTasksChoice -ne "0") { Pause }
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
                        "8" { Remove-HPBloatware }
                        "0" { break }
                        default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
                    }
                    if ($newPCSetupChoice -ne "0") { Pause }
                } while ($newPCSetupChoice -ne "0")
            }
            "0" { Write-Host "Exiting..." -ForegroundColor Yellow }
            default { Write-Host "Invalid choice, please try again." -ForegroundColor Red }
        }
        if ($mainChoice -notmatch "^(0|[1-4](\.[1-9]+)?)$") { Pause }
    } while ($mainChoice -ne "0")
}
