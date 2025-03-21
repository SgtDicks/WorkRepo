# ===============================
# RPI Repair and Maintenance Tool
# ===============================

# Set Execution Policy and load required assemblies for GUI
Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# ===============================
# Helper and Utility Functions
# ===============================

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

# Function to repair Windows using DISM
function Repair-Windows {
    Write-Host "Running DISM command to restore health..." -ForegroundColor Green
    Write-Host "This process can take 15-30 minutes and may require a restart." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        dism /online /cleanup-image /restorehealth
        Write-Host "Windows repair completed." -ForegroundColor Green
    }
}

# Function to repair system files using SFC
function Repair-SystemFiles {
    Write-Host "Running System File Checker (SFC)..." -ForegroundColor Green
    Write-Host "This scan can take up to 20 minutes." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        sfc /scannow
        Write-Host "System File Checker completed." -ForegroundColor Green
    }
}

# Function to repair disk using CHKDSK
function Repair-Disk {
    Write-Host "Running Check Disk (CHKDSK)..." -ForegroundColor Green
    Write-Host "This operation may take several hours and will restart the computer." -ForegroundColor Red
    if (Confirm-Action "This operation may take several hours and will restart the computer.") {
        chkdsk C: /F /R /X
        Write-Host "Check Disk completed." -ForegroundColor Green
    }
}

# Function to run Windows Update Troubleshooter
function Run-WindowsUpdateTroubleshooter {
    Write-Host "Running Windows Update Troubleshooter..." -ForegroundColor Green
    Write-Host "This may take 5-10 minutes." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time.") {
        Start-Process -FilePath "msdt.exe" -ArgumentList "/id WindowsUpdateDiagnostic" -Wait
        Write-Host "Windows Update Troubleshooter completed." -ForegroundColor Green
    }
}

# Function to check and repair DISM
function Check-And-Repair-DISM {
    Write-Host "Running DISM Check and Repair..." -ForegroundColor Green
    Write-Host "This process may take 30-60 minutes and a restart might be required." -ForegroundColor Yellow
    if (Confirm-Action "This operation may take some time and require a restart.") {
        DISM /Online /Cleanup-Image /CheckHealth
        DISM /Online /Cleanup-Image /ScanHealth
        DISM /Online /Cleanup-Image /RestoreHealth
        Write-Host "DISM Check and Repair completed." -ForegroundColor Green
    }
}

# Function to reset network
function Reset-Network {
    Write-Host "Resetting Network Adapters..." -ForegroundColor Green
    Write-Host "This will cause a temporary network outage." -ForegroundColor Red
    if (Confirm-Action "This will temporarily disrupt network connectivity.") {
        netsh winsock reset
        netsh int ip reset
        ipconfig /release
        ipconfig /renew
        ipconfig /flushdns
        Write-Host "Network reset completed." -ForegroundColor Green
    }
}

# Function to run Windows Memory Diagnostic
function Run-MemoryDiagnostic {
    Write-Host "Running Windows Memory Diagnostic..." -ForegroundColor Green
    Write-Host "This test will restart your computer." -ForegroundColor Red
    if (Confirm-Action "This operation will restart your computer.") {
        Start-Process -FilePath "mdsched.exe" -ArgumentList "/f" -Verb RunAs
        Write-Host "Memory Diagnostic scheduled." -ForegroundColor Green
    }
}

# Function to run Startup Repair
function Run-StartupRepair {
    Write-Host "Running Startup Repair..." -ForegroundColor Green
    Write-Host "This process will restart your computer." -ForegroundColor Red
    if (Confirm-Action "This operation will restart your computer.") {
        Start-Process -FilePath "reagentc.exe" -ArgumentList "/boottore" -Verb RunAs -Wait
        shutdown /r /t 0
    }
}

# Function to run Windows Defender Full Scan
function Run-WindowsDefenderScan {
    Write-Host "Running Windows Defender Full Scan..." -ForegroundColor Green
    Write-Host "This scan may take several hours." -ForegroundColor Yellow
    if (Confirm-Action "This scan may take several hours.") {
        Start-MpScan -ScanType FullScan
        Write-Host "Windows Defender Full Scan completed." -ForegroundColor Green
    }
}

# Function to reset Windows Update components
function Reset-WindowsUpdateComponents {
    Write-Host "Resetting Windows Update components..." -ForegroundColor Green
    Write-Host "This may take 10-20 minutes." -ForegroundColor Yellow
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

# Function to clear Teams cache (updated to handle wildcards)
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
        # Use Get-ChildItem to process wildcard paths
        $teamsCacheDirs = Get-ChildItem -Path "$env:LOCALAPPDATA\Packages" -Filter "MSTeams_*" -Directory -ErrorAction SilentlyContinue
        $cacheRemoved = $false
        foreach ($dir in $teamsCacheDirs) {
            $cachePath = Join-Path $dir.FullName "LocalCache\Local\Microsoft\Teams"
            if (Test-Path $cachePath) {
                try {
                    Remove-Item -Path $cachePath -Recurse -Force -ErrorAction Stop
                    Write-Host "Teams cache removed from: $cachePath" -ForegroundColor Green
                    $cacheRemoved = $true
                } catch {
                    Write-Warning "Failed to remove Teams cache at $cachePath. $_"
                }
            } else {
                Write-Host "Teams cache path not found: $cachePath" -ForegroundColor Yellow
            }
        }
        if (-not $cacheRemoved) {
            Write-Warning "No Teams cache directories were found or removed."
        }
        Write-Host "Cleanup complete... Trying to launch Teams" -ForegroundColor Green
        Start-Teams
    } else {
        Write-Host "Cache deletion canceled." -ForegroundColor Yellow
    }
}

# Function to display the Windows Repairs submenu (Console Mode)
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
        Write-Host "List of installed applications displayed." -ForegroundColor Green
    } catch {
        Write-Warning "Failed to list installed applications. $_"
    }
}

# ===============================
# Network Diagnostics Function
# ===============================

function Network-Diagnostics {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Network Diagnostics Results"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)

    $resultsTextBox = New-Object System.Windows.Forms.RichTextBox
    $resultsTextBox.Size = New-Object System.Drawing.Size(760, 500)
    $resultsTextBox.Location = New-Object System.Drawing.Point(20,20)
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.BackColor = [System.Drawing.Color]::White
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas",10)
    $form.Controls.Add($resultsTextBox)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Size = New-Object System.Drawing.Size(100,30)
    $closeButton.Location = New-Object System.Drawing.Point(680,530)
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export Results"
    $exportButton.Size = New-Object System.Drawing.Size(120,30)
    $exportButton.Location = New-Object System.Drawing.Point(550,530)
    $exportButton.Add_Click({
        param($sender, $e)
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        $saveDialog.Title = "Save Network Diagnostics Results"
        $saveDialog.FileName = "NetworkDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        if ($saveDialog.ShowDialog() -eq 'OK') {
            $resultsTextBox.Text | Out-File -FilePath $saveDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("Results exported to $($saveDialog.FileName)", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $form.Controls.Add($exportButton)

    function Add-ColoredText {
        param(
            [string]$text,
            [System.Drawing.Color]$color = [System.Drawing.Color]::Black
        )
        $resultsTextBox.SelectionStart = $resultsTextBox.TextLength
        $resultsTextBox.SelectionLength = 0
        $resultsTextBox.SelectionColor = $color
        $resultsTextBox.AppendText($text)
        $resultsTextBox.SelectionColor = $resultsTextBox.ForeColor
    }

    $resultsTextBox.Clear()
    Add-ColoredText "===== NETWORK ADAPTER INFORMATION =====`r`n" -color [System.Drawing.Color]::Blue
    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        foreach ($adapter in $adapters) {
            Add-ColoredText "Adapter: $($adapter.Name) ($($adapter.InterfaceDescription))`r`n" -color [System.Drawing.Color]::DarkBlue
            Add-ColoredText "Status: $($adapter.Status)`r`n" -color [System.Drawing.Color]::Green
            Add-ColoredText "MAC Address: $($adapter.MacAddress)`r`n" -color [System.Drawing.Color]::Black
            Add-ColoredText "Link Speed: $($adapter.LinkSpeed)`r`n`r`n" -color [System.Drawing.Color]::Black
        }
    } catch {
        Add-ColoredText "Error retrieving network adapter information: $_`r`n" -color [System.Drawing.Color]::Red
    }
    Add-ColoredText "===== IP CONFIGURATION =====`r`n" -color [System.Drawing.Color]::Blue
    try {
        $ipConfig = Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq "Up" }
        foreach ($config in $ipConfig) {
            Add-ColoredText "Interface: $($config.InterfaceAlias)`r`n" -color [System.Drawing.Color]::DarkBlue
            $ipv4 = $config | Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue
            if ($ipv4) {
                Add-ColoredText "IPv4 Address: $($ipv4.IPAddress)`r`n" -color [System.Drawing.Color]::Black
                Add-ColoredText "Subnet Mask: $($ipv4.PrefixLength)`r`n" -color [System.Drawing.Color]::Black
            }
            if ($config.IPv4DefaultGateway) {
                Add-ColoredText "Default Gateway: $($config.IPv4DefaultGateway.NextHop)`r`n" -color [System.Drawing.Color]::Black
            } else {
                Add-ColoredText "Default Gateway: Not configured`r`n" -color [System.Drawing.Color]::Red
            }
            if ($config.DNSServer) {
                Add-ColoredText "DNS Servers: $($config.DNSServer.ServerAddresses -join ', ')`r`n" -color [System.Drawing.Color]::Black
            } else {
                Add-ColoredText "DNS Servers: Not configured`r`n" -color [System.Drawing.Color]::Red
            }
            Add-ColoredText "`r`n"
        }
    } catch {
        Add-ColoredText "Error retrieving IP configuration: $_`r`n" -color [System.Drawing.Color]::Red
    }
    # (Additional diagnostic tests follow here...)

    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()
}

# ===============================
# Power Management Functions
# ===============================

function Check-BatteryHealth {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Battery Health Report"
    $form.Size = New-Object System.Drawing.Size(700,500)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)

    $resultsTextBox = New-Object System.Windows.Forms.RichTextBox
    $resultsTextBox.Size = New-Object System.Drawing.Size(660,400)
    $resultsTextBox.Location = New-Object System.Drawing.Point(20,20)
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.BackColor = [System.Drawing.Color]::White
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas",10)
    $form.Controls.Add($resultsTextBox)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Size = New-Object System.Drawing.Size(100,30)
    $closeButton.Location = New-Object System.Drawing.Point(580,430)
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Generate Full Report"
    $exportButton.Size = New-Object System.Drawing.Size(150,30)
    $exportButton.Location = New-Object System.Drawing.Point(420,430)
    $exportButton.Add_Click({
        param($sender, $e)
        $reportPath = "$env:USERPROFILE\battery-report.html"
        powercfg /batteryreport /output $reportPath
        if (Test-Path $reportPath) {
            Start-Process $reportPath
            [System.Windows.Forms.MessageBox]::Show("Full battery report generated at $reportPath", "Report Generated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to generate battery report", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($exportButton)

    function Add-ColoredText {
        param(
            [string]$text,
            [System.Drawing.Color]$color = [System.Drawing.Color]::Black
        )
        $resultsTextBox.SelectionStart = $resultsTextBox.TextLength
        $resultsTextBox.SelectionLength = 0
        $resultsTextBox.SelectionColor = $color
        $resultsTextBox.AppendText($text)
        $resultsTextBox.SelectionColor = $resultsTextBox.ForeColor
    }

    $resultsTextBox.Clear()
    Add-ColoredText "===== BATTERY HEALTH CHECK =====`r`n" -color [System.Drawing.Color]::Blue
    try {
        $batteryInfo = Get-WmiObject -Class Win32_Battery
        if ($batteryInfo) {
            Add-ColoredText "Battery Found: Yes`r`n" -color [System.Drawing.Color]::Green
            Add-ColoredText "Battery Name: $($batteryInfo.Name)`r`n" -color [System.Drawing.Color]::Black
            Add-ColoredText "Description: $($batteryInfo.Description)`r`n" -color [System.Drawing.Color]::Black
            Add-ColoredText "Battery Status: " -color [System.Drawing.Color]::Black
            switch ($batteryInfo.BatteryStatus) {
                1 { Add-ColoredText "Discharging`r`n" -color [System.Drawing.Color]::Orange }
                2 { Add-ColoredText "AC Power`r`n" -color [System.Drawing.Color]::Green }
                3 { Add-ColoredText "Fully Charged`r`n" -color [System.Drawing.Color]::Green }
                4 { Add-ColoredText "Low`r`n" -color [System.Drawing.Color]::Red }
                5 { Add-ColoredText "Critical`r`n" -color [System.Drawing.Color]::Red }
                default { Add-ColoredText "Unknown ($($batteryInfo.BatteryStatus))`r`n" -color [System.Drawing.Color]::Gray }
            }
            $chargePercent = $batteryInfo.EstimatedChargeRemaining
            Add-ColoredText "Current Charge: $chargePercent%`r`n" -color [System.Drawing.Color]::Black
            # (Additional capacity analysis code would be here)
        }
        else {
            Add-ColoredText "No battery detected. This device appears to be a desktop or server without a battery.`r`n" -color [System.Drawing.Color]::Red
        }
    }
    catch {
        Add-ColoredText "Error checking battery status: $_`r`n" -color [System.Drawing.Color]::Red
    }
    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()
}

# Function to manage power plans
function Manage-PowerPlans {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Power Plan Management"
    $form.Size = New-Object System.Drawing.Size(600,450)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Size = New-Object System.Drawing.Size(560,250)
    $listView.Location = New-Object System.Drawing.Point(20,20)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Columns.Add("GUID",280) | Out-Null
    $listView.Columns.Add("Name",200) | Out-Null
    $listView.Columns.Add("Active",60) | Out-Null
    $form.Controls.Add($listView)

    $activateButton = New-Object System.Windows.Forms.Button
    $activateButton.Text = "Activate Selected Plan"
    $activateButton.Size = New-Object System.Drawing.Size(150,30)
    $activateButton.Location = New-Object System.Drawing.Point(20,280)
    $form.Controls.Add($activateButton)

    $balancedButton = New-Object System.Windows.Forms.Button
    $balancedButton.Text = "Balanced Plan"
    $balancedButton.Size = New-Object System.Drawing.Size(120,30)
    $balancedButton.Location = New-Object System.Drawing.Point(20,320)
    $form.Controls.Add($balancedButton)

    $powerSaverButton = New-Object System.Windows.Forms.Button
    $powerSaverButton.Text = "Power Saver"
    $powerSaverButton.Size = New-Object System.Drawing.Size(120,30)
    $powerSaverButton.Location = New-Object System.Drawing.Point(150,320)
    $form.Controls.Add($powerSaverButton)

    $highPerfButton = New-Object System.Windows.Forms.Button
    $highPerfButton.Text = "High Performance"
    $highPerfButton.Size = New-Object System.Drawing.Size(120,30)
    $highPerfButton.Location = New-Object System.Drawing.Point(280,320)
    $form.Controls.Add($highPerfButton)

    $ultimatePerfButton = New-Object System.Windows.Forms.Button
    $ultimatePerfButton.Text = "Ultimate Performance"
    $ultimatePerfButton.Size = New-Object System.Drawing.Size(140,30)
    $ultimatePerfButton.Location = New-Object System.Drawing.Point(410,320)
    $form.Controls.Add($ultimatePerfButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Size = New-Object System.Drawing.Size(100,30)
    $closeButton.Location = New-Object System.Drawing.Point(480,380)
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Size = New-Object System.Drawing.Size(450,20)
    $statusLabel.Location = New-Object System.Drawing.Point(20,380)
    $statusLabel.Text = "Select a power plan to activate"
    $form.Controls.Add($statusLabel)

    function Refresh-PowerPlanList {
        $listView.Items.Clear()
        $powerPlans = powercfg /list | Where-Object { $_ -match "Power Scheme GUID:" }
        $activePlanGuid = (powercfg /getactivescheme) -replace '.*GUID: ([a-z0-9-]+).*', '$1'
        foreach ($plan in $powerPlans) {
            $planGuid = $plan -replace '.*GUID: ([a-z0-9-]+).*', '$1'
            $planName = $plan -replace '.*\((.*)\).*', '$1'
            $isActive = if ($planGuid -eq $activePlanGuid) { "Yes" } else { "No" }
            $item = New-Object System.Windows.Forms.ListViewItem($planGuid)
            $item.SubItems.Add($planName)
            $item.SubItems.Add($isActive)
            if ($isActive -eq "Yes") {
                $item.BackColor = [System.Drawing.Color]::LightGreen
            }
            $listView.Items.Add($item)
        }
    }

    $activateButton.Add_Click({
        param($sender, $e)
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedGuid = $listView.SelectedItems[0].Text
            powercfg /setactive $selectedGuid
            $statusLabel.Text = "Power plan activated: $($listView.SelectedItems[0].SubItems[1].Text)"
            Refresh-PowerPlanList
        } else {
            $statusLabel.Text = "Please select a power plan first"
        }
    })
    $balancedButton.Add_Click({
        param($sender, $e)
        powercfg /setactive SCHEME_BALANCED
        $statusLabel.Text = "Balanced power plan activated"
        Refresh-PowerPlanList
    })
    $powerSaverButton.Add_Click({
        param($sender, $e)
        powercfg /setactive SCHEME_MAX_BATTERY_LIFE
        $statusLabel.Text = "Power saver plan activated"
        Refresh-PowerPlanList
    })
    $highPerfButton.Add_Click({
        param($sender, $e)
        powercfg /setactive SCHEME_MIN_POWER_SAVINGS
        $statusLabel.Text = "High performance plan activated"
        Refresh-PowerPlanList
    })
    $ultimatePerfButton.Add_Click({
        param($sender, $e)
        $ultimateExists = powercfg /list | Where-Object { $_ -match "Ultimate Performance" }
        if (-not $ultimateExists) {
            powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61
            $statusLabel.Text = "Ultimate Performance plan created and activated"
        } else {
            $ultimateGuid = ($ultimateExists -replace '.*GUID: ([a-z0-9-]+).*', '$1')
            powercfg /setactive $ultimateGuid
            $statusLabel.Text = "Ultimate Performance plan activated"
        }
        Refresh-PowerPlanList
    })

    $form.Add_Shown({
        Refresh-PowerPlanList
        $form.Activate()
    })
    [void]$form.ShowDialog()
}

# Function to display the Power Management submenu GUI (new function)
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
    $batteryButton.Add_Click({
        param($sender, $e)
        Check-BatteryHealth
    })
    $powerForm.Controls.Add($batteryButton)
    
    $powerPlanButton = New-Object System.Windows.Forms.Button
    $powerPlanButton.Text = "Manage Power Plans"
    $powerPlanButton.Size = New-Object System.Drawing.Size(200,40)
    $powerPlanButton.Location = New-Object System.Drawing.Point(100,120)
    $powerPlanButton.BackColor = [System.Drawing.Color]::LightGreen
    $powerPlanButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $powerPlanButton.Add_Click({
        param($sender, $e)
        Manage-PowerPlans
    })
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

# (Additional functions such as Manage-SleepSettings, Generate-BatteryReport, Run-PowerTroubleshooter,
# Clean-TempFiles, Check-DiskUsage, Show-OfficeMenu, Show-UserTasksMenu, Show-PrinterMenu,
# Map-Printer, Open-NewPCFiles, Download-And-Open-Ninite, Download-MS-Teams, Install-AdobeReader,
# Remove-HPBloatware, Install-AllPrinters, and Show-NewPCSetupMenu would follow here.)

# ===============================
# GUI Main Menu Functions
# ===============================

function Show-WindowsMenuGUI {
    $windowsForm = New-Object System.Windows.Forms.Form
    $windowsForm.Text = "Windows Repairs"
    $windowsForm.Size = New-Object System.Drawing.Size(600,500)
    $windowsForm.StartPosition = "CenterScreen"
    $windowsForm.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $windowsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $windowsForm.MaximizeBox = $false

    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Windows Repair Options"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
    $titleLabel.Size = New-Object System.Drawing.Size(560,30)
    $titleLabel.Location = New-Object System.Drawing.Point(20,20)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $windowsForm.Controls.Add($titleLabel)
    
    # Create buttons for Windows repair options
    $buttonWidth = 170
    $buttonHeight = 60
    $buttonsPerRow = 3
    $horizontalSpacing = 15
    $verticalSpacing = 10
    $startX = 20
    $startY = 60
    $windowsOptions = @(
        @{ Text = "DISM RestoreHealth"; Action = { Repair-Windows } },
        @{ Text = "System File Checker"; Action = { Repair-SystemFiles } },
        @{ Text = "Check Disk"; Action = { Repair-Disk } },
        @{ Text = "Windows Update Troubleshooter"; Action = { Run-WindowsUpdateTroubleshooter } },
        @{ Text = "DISM Check and Repair"; Action = { Check-And-Repair-DISM } },
        @{ Text = "Network Reset"; Action = { Reset-Network } },
        @{ Text = "Memory Diagnostic"; Action = { Run-MemoryDiagnostic } },
        @{ Text = "Startup Repair"; Action = { Run-StartupRepair } },
        @{ Text = "Windows Defender Scan"; Action = { Run-WindowsDefenderScan } },
        @{ Text = "Reset Update Components"; Action = { Reset-WindowsUpdateComponents } },
        @{ Text = "List Installed Apps"; Action = { List-InstalledApps } },
        @{ Text = "Factory Reset"; Action = { Factory-Reset } }
    )
    
    $currentX = $startX
    $currentY = $startY
    $buttonCount = 0
    foreach ($option in $windowsOptions) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $option.Text
        $button.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
        $button.Location = New-Object System.Drawing.Point($currentX,$currentY)
        $button.BackColor = [System.Drawing.Color]::LightBlue
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.Tag = $option.Action
        $button.Add_Click({
            param($sender, $e)
            $scriptBlock = $sender.Tag
            & $scriptBlock
        })
        $windowsForm.Controls.Add($button)
        $buttonCount++
        if ($buttonCount % $buttonsPerRow -eq 0) {
            $currentX = $startX
            $currentY += $buttonHeight + $verticalSpacing
        } else {
            $currentX += $buttonWidth + $horizontalSpacing
        }
    }
    
    # Back button
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

    $diskUsageButton = New-Object System.Windows.Forms.Button
    $diskUsageButton.Text = "Check Disk Usage"
    $diskUsageButton.Size = New-Object System.Drawing.Size(200,40)
    $diskUsageButton.Location = New-Object System.Drawing.Point(100,220)
    $diskUsageButton.BackColor = [System.Drawing.Color]::LightYellow
    $diskUsageButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $diskUsageButton.Add_Click({ Check-DiskUsage })
    $userTasksForm.Controls.Add($diskUsageButton)

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

# Function to display the Printer Mapping submenu
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
        $button.Add_Click({
            param($sender, $e)
            $printerIP = $sender.Tag
            Map-Printer -PrinterIP $printerIP
        })
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
    Write-Host "0: Back to Main Menu"
}

# ===============================
# Main Menu GUI
# ===============================

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

    # Second row
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

    # Power Management button (now calls our new function)
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

# ===============================
# Main Script Loop (Console Fallback)
# ===============================
$guiMode = $true
try {
    [System.Windows.Forms.Application] | Out-Null
} catch {
    $guiMode = $false
}

if ($guiMode) {
    Show-MainMenuGUI
} else {
    Write-Host "GUI mode not available. Falling back to console mode." -ForegroundColor Yellow
    # (Console mode code would be implemented here)
}
