#Requires -Version 5.1
#Requires -RunAsAdministrator

# --- INITIAL SETUP (Pre-GUI, logs to console, then GUI will recap) ---
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser -ErrorAction SilentlyContinue
    if ($currentPolicy -ne "RemoteSigned" -and $currentPolicy -ne "Unrestricted") {
        Write-Host "Current Execution Policy for CurrentUser is $currentPolicy. Setting to RemoteSigned."
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "Execution Policy for CurrentUser set to RemoteSigned." -ForegroundColor Green
    } else {
        Write-Host "Execution Policy for CurrentUser is already sufficient ($currentPolicy)." -ForegroundColor Yellow
    }
} catch {
    Write-Warning "Failed to set Execution Policy. $_. Some script features might not work."
}

$global:appsPath = "C:\apps" # Make it globally accessible for functions
if (-not (Test-Path -Path $global:appsPath)) {
    try {
        New-Item -ItemType Directory -Path $global:appsPath -ErrorAction Stop | Out-Null
        Write-Host "Created directory: $($global:appsPath)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create directory $($global:appsPath). $_. Some downloads may fail."
    }
} else {
    Write-Host "Directory already exists: $($global:appsPath)" -ForegroundColor Yellow
}

# --- GUI SETUP ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global Output RichTextBox
$script:outputBox = New-Object System.Windows.Forms.RichTextBox
$script:outputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:outputBox.ReadOnly = $true
$script:outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$script:outputBox.WordWrap = $true
$script:outputBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$script:outputBox.HideSelection = $false

# Helper function to write to the GUI output box
function Write-GuiLog {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = ([System.Drawing.Color]::FromName('Black')),
        [switch]$NoTimestamp
    )
    
    if ($script:outputBox.IsDisposed) { return }

    $script:outputBox.Invoke([Action]{
        $timestamp = if ($NoTimestamp) { "" } else { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - " }
        $script:outputBox.SelectionStart = $script:outputBox.TextLength
        $script:outputBox.SelectionLength = 0
        $script:outputBox.SelectionColor = $Color
        $script:outputBox.AppendText("$timestamp$Message`r`n")
        $script:outputBox.ScrollToCaret()
    })
}

# --- MODIFIED FUNCTION DEFINITIONS ---

function Start-LongRunningJob {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        [Parameter(Mandatory=$true)]
        [string]$OperationName,
        [System.Windows.Forms.Button]$ButtonToDisable
    )

    if ($ButtonToDisable -and -not $ButtonToDisable.IsDisposed) { $ButtonToDisable.Enabled = $false }
    Write-GuiLog "Starting background job: $OperationName..." -Color ([System.Drawing.Color]::FromArgb(255, 100, 100, 255))

    $job = Start-Job -ScriptBlock $ScriptBlock -Name $OperationName
    
    Register-ObjectEvent -InputObject $job -EventName StateChanged -SourceIdentifier "JobEvent_$($job.Id)_$($OperationName.Replace(' ','_'))" -Action {
        $evtJob = $Sender
        $jobState = $evtJob.JobStateInfo.State
        $sourceId = $EventArgs.SourceIdentifier
        
        $script:outputBox.Invoke([Action]{
            Write-GuiLog "Job '$($evtJob.Name)' state: $jobState" -Color Gray
            if ($jobState -in ('Completed', 'Failed', 'Stopped')) {
                $jobErrors = $evtJob.ChildJobs[0].Error
                if ($jobErrors.Count -gt 0) {
                    Write-GuiLog "Errors from job '$($evtJob.Name)':" -Color Red
                    $jobErrors | ForEach-Object { Write-GuiLog $_.ToString() -Color Red -NoTimestamp }
                }
                
                $output = Receive-Job -Job $evtJob -Keep
                if ($output) {
                    Write-GuiLog "Output from job '$($evtJob.Name)':" -Color DarkGray
                    $output | ForEach-Object { Write-GuiLog $_.ToString() -Color DarkGray -NoTimestamp }
                }

                if ($jobState -eq 'Completed') {
                     Write-GuiLog "Job '$($evtJob.Name)' completed successfully." -Color Green
                } else {
                     Write-GuiLog "Job '$($evtJob.Name)' finished with state: $jobState. Reason: $($evtJob.JobStateInfo.Reason)" -Color Red
                }

                if ($ButtonToDisable -and -not $ButtonToDisable.IsDisposed) { $ButtonToDisable.Enabled = $true }
                
                try { Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue } catch {}
                try { Remove-Job -Job $evtJob -ErrorAction SilentlyContinue } catch {}
            }
        })
    } | Out-Null
    Write-GuiLog "Job '$OperationName' (ID: $($job.Id)) submitted. See log for updates." -Color ([System.Drawing.Color]::FromArgb(255,100,100,255))
}

function Change-PCName-Action {
    Write-GuiLog "Changing the PC name based on the serial number..." -Color Cyan
    try {
        $serialNumber = (Get-WmiObject -Class Win32_BIOS -ErrorAction Stop).SerialNumber
        $newPCName = "RPI-" + $serialNumber.Trim()
        Write-GuiLog "The new PC name will be: $newPCName" -Color Yellow
        Write-GuiLog "Computer will be renamed and RESTART. Ensure all work is saved." -Color Red
        Rename-Computer -NewName $newPCName -Force -Restart
        Write-GuiLog "Rename command issued. System should restart shortly." -Color Green
    } catch {
        Write-GuiLog "Error changing PC name: $($_.Exception.Message)" -Color Red
    }
}

function Join-Domain-Action {
    param (
        [string]$domainName,
        [string]$ouPath,
        [System.Management.Automation.PSCredential]$credential
    )
    Write-GuiLog "Joining the computer to the domain $domainName..." -Color Cyan
    if ($ouPath) {
        Write-GuiLog "Target OU Path: $ouPath" -Color Cyan
    } else {
        Write-GuiLog "No OU Path specified, using default computer container." -Color Yellow
    }
    Write-GuiLog "Computer will be joined to domain and RESTART. Ensure all work is saved." -Color Red
    
    $commandParams = @{
        DomainName = $domainName
        Credential = $credential
        Force      = $true
        Restart    = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($ouPath)) {
        $commandParams.OUPath = $ouPath
    }

    try {
        Add-Computer @commandParams -ErrorAction Stop
        Write-GuiLog "Join domain command issued. System should restart shortly." -Color Green
    } catch {
        if ($_.Exception -is [System.InvalidOperationException] -and ($_.Exception.Message -like "*already in that domain*" -or $_.Exception.Message -like "*already a member of domain*")) {
            Write-GuiLog "Computer is already a member of the domain '$domainName'." -Color Yellow
        } else {
            Write-GuiLog "Error joining domain '$domainName': $($_.Exception.Message)" -Color Red
        }
    }
}

function Repair-Windows-ScriptBlock {
    Write-GuiLog "Running DISM command to restore health..." -Color Green
    Write-GuiLog "This process can take 15-30 minutes. A restart may be required." -Color Yellow
    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting DISM /Online /Cleanup-Image /RestoreHealth..."
        dism /online /cleanup-image /restorehealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows repair (DISM RestoreHealth) completed."
    }
}

function Repair-SystemFiles-ScriptBlock {
    Write-GuiLog "Running System File Checker (SFC)..." -Color Green
    Write-GuiLog "This scan can take up to 20 minutes. No restart is required unless issues are found." -Color Yellow
    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting SFC /scannow..."
        sfc /scannow
        Microsoft.PowerShell.Host.WriteTranscriptUtil "System File Checker (SFC) completed."
    }
}

function Repair-Disk-Action {
    Write-GuiLog "Running Check Disk (CHKDSK) on C:..." -Color Green
    Write-GuiLog "This operation could take a few hours and WILL RESTART the computer. Ensure all work is saved." -Color Red
    try {
        chkdsk C: /F /R /X
        Write-GuiLog "CHKDSK scheduled for the next restart. The system might prompt for restart or restart automatically." -Color Green
    } catch {
        Write-GuiLog "Error scheduling CHKDSK: $($_.Exception.Message)" -Color Red
    }
}

function Run-WindowsUpdateTroubleshooter-Action {
    Write-GuiLog "Running Windows Update Troubleshooter..." -Color Green
    Write-GuiLog "This operation might take 5-10 minutes. An interactive window will open." -Color Yellow
    try {
        Start-Process -FilePath "msdt.exe" -ArgumentList "/id WindowsUpdateDiagnostic"
        Write-GuiLog "Windows Update Troubleshooter started. Please follow its prompts." -Color Green
    } catch {
         Write-GuiLog "Failed to start Windows Update Troubleshooter: $($_.Exception.Message)" -Color Red
    }
}

function Check-And-Repair-DISM-ScriptBlock {
    Write-GuiLog "Running DISM Check and Repair..." -Color Green
    Write-GuiLog "This process might take 30-60 minutes. A restart may be required." -Color Yellow
    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running DISM /Online /Cleanup-Image /CheckHealth..."
        DISM /Online /Cleanup-Image /CheckHealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running DISM /Online /Cleanup-Image /ScanHealth..."
        DISM /Online /Cleanup-Image /ScanHealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running DISM /Online /Cleanup-Image /RestoreHealth..."
        DISM /Online /Cleanup-Image /RestoreHealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "DISM Check and Repair completed."
    }
}

function Reset-Network-ScriptBlock {
    Write-GuiLog "Resetting Network Adapters..." -Color Green
    Write-GuiLog "This will cause a temporary network outage and may require a restart." -Color Red
    return {
        $ErrorActionPreference = 'Stop'
        $Commands = @(
            { netsh winsock reset },
            { netsh int ip reset },
            { ipconfig /release },
            { ipconfig /renew },
            { ipconfig /flushdns }
        )
        foreach($cmd in $Commands){
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Executing: $($cmd.ToString())"
            Invoke-Command -ScriptBlock $cmd
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Network reset commands executed. A restart may be required."
    }
}

function Run-MemoryDiagnostic-Action {
    Write-GuiLog "Scheduling Windows Memory Diagnostic..." -Color Green
    Write-GuiLog "This test WILL RESTART your computer. Ensure all work is saved." -Color Red
    try {
        Start-Process -FilePath "mdsched.exe" -ArgumentList "/f" -Verb RunAs
        Write-GuiLog "Windows Memory Diagnostic scheduled. The computer will restart to run the test." -Color Green
    } catch {
        Write-GuiLog "Failed to schedule Windows Memory Diagnostic: $($_.Exception.Message)" -Color Red
    }
}

function Run-StartupRepair-Action {
    Write-GuiLog "Initiating Startup Repair..." -Color Green
    Write-GuiLog "This process WILL RESTART your computer and attempt to fix startup issues." -Color Red
    try {
        Write-GuiLog "Configuring boot to recovery environment..." -Color Cyan
        Start-Process -FilePath "reagentc.exe" -ArgumentList "/boottore" -Verb RunAs -Wait
        Write-GuiLog "Boot to recovery configured. Restarting computer NOW..." -Color Cyan
        Shutdown.exe /r /t 0 /f
    } catch {
        Write-GuiLog "Failed to initiate Startup Repair: $($_.Exception.Message)" -Color Red
    }
}

function Run-WindowsDefenderScan-ScriptBlock {
    Write-GuiLog "Running Windows Defender Full Scan..." -Color Green
    Write-GuiLog "This scan can take several hours. No restart is required." -Color Yellow
    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting Windows Defender Full Scan. This can take a long time..."
        Start-MpScan -ScanType FullScan -ErrorAction Stop
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows Defender Full Scan process completed or handed off. Monitor Windows Security for progress."
    }
}

function Reset-WindowsUpdateComponents-ScriptBlock {
    Write-GuiLog "Resetting Windows Update components..." -Color Green
    Write-GuiLog "This operation may take 10-20 minutes. Services will be temporarily unavailable." -Color Yellow
    return {
        $ErrorActionPreference = 'Stop'
        $servicesToManage = @("wuauserv", "cryptSvc", "bits", "msiserver")
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Stopping services: $($servicesToManage -join ', ')"
        Stop-Service -Name $servicesToManage -Force -ErrorAction SilentlyContinue

        $pathsToRename = @{
            "C:\Windows\SoftwareDistribution" = "SoftwareDistribution.old"
            "C:\Windows\System32\catroot2"   = "catroot2.old"
        }
        foreach ($entry in $pathsToRename.GetEnumerator()) {
            $oldPath = $entry.Key
            $newDirName = $entry.Value
            $parentDir = Split-Path $oldPath
            $newFullPath = Join-Path $parentDir $newDirName
            
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Attempting to rename $oldPath to $newFullPath"
            if (Test-Path $newFullPath) {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing existing $newFullPath..."
                Remove-Item -Path $newFullPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path $oldPath) {
                Rename-Item -Path $oldPath -NewName $newDirName -Force -ErrorAction Stop
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Renamed $oldPath to $newDirName"
            } else {
                 Microsoft.PowerShell.Host.WriteTranscriptUtil "$oldPath not found."
            }
        }
        
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting services: $($servicesToManage -join ', ')"
        Start-Service -Name $servicesToManage -ErrorAction SilentlyContinue
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows Update components reset completed."
    }
}

function Start-Teams-Action {
    Write-GuiLog "Trying to find and launch Microsoft Teams..." -Color Cyan
    $possiblePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
        (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\ms-teams.exe"),
        "$env:LOCALAPPDATA\Programs\Teams\current\Teams.exe",
        "$env:PROGRAMFILES\Teams Installer\Teams.exe",
        "$env:PROGRAMFILES(X86)\Teams Installer\Teams.exe"
    )
    $teamsPath = $null
    foreach ($path in $possiblePaths) {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
        if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
            $teamsPath = $resolvedPath
            Write-GuiLog "Teams executable found at: $teamsPath" -Color Green
            break
        } else {
            Write-GuiLog "Teams not found at: $resolvedPath (Skipping)" -Color DarkGray
        }
    }

    if ($teamsPath) {
        try {
            Start-Process -FilePath $teamsPath
            Write-GuiLog "Microsoft Teams launched." -Color Green
        } catch {
            Write-GuiLog "Failed to start Teams from ${teamsPath}: $($_.Exception.Message)" -Color Red
        }
    } else {
        Write-GuiLog "Microsoft Teams executable not found in common locations." -Color Red
    }
}

function Clear-TeamsCache-ScriptBlock {
    Write-GuiLog "Attempting to clear Microsoft Teams cache..." -Color Cyan
    return {
        $ErrorActionPreference = 'Continue'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Closing Teams processes (ms-teams, Teams)..."
        Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3

        $cacheLocations = @(
            "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe",
            "$env:LOCALAPPDATA\Packages\MicrosoftTeams_8wekyb3d8bbwe",
            "$env:LOCALAPPDATA\Microsoft\Teams",
            "$env:APPDATA\Microsoft\Teams"
        )
        
        foreach ($location in $cacheLocations) {
            $resolvedLocation = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($location)
            if (Test-Path -LiteralPath $resolvedLocation) {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Attempting to remove: $resolvedLocation"
                try {
                    Remove-Item -Path $resolvedLocation -Recurse -Force -ErrorAction Stop
                    Microsoft.PowerShell.Host.WriteTranscriptUtil "Successfully removed: $resolvedLocation"
                } catch {
                    Microsoft.PowerShell.Host.WriteTranscriptUtil "WARN: Failed to remove $resolvedLocation. $($_.Exception.Message)"
                }
            } else {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Path not found (or already removed): $resolvedLocation"
            }
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Teams cache clearing process complete. Attempting to restart Teams..."
        
        $possiblePathsToStart = @(
            "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
            (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\ms-teams.exe")
        )
        $teamsExeToStart = $null
        foreach ($p in $possiblePathsToStart) {
            $rp = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
            if (Test-Path -LiteralPath $rp -PathType Leaf) { $teamsExeToStart = $rp; break }
        }
        if ($teamsExeToStart) {
            Start-Process -FilePath $teamsExeToStart
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Attempted to start Teams from $teamsExeToStart."
        } else {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Could not find Teams executable to auto-restart. Please start manually."
        }
    }
}

function List-InstalledApps-ScriptBlock {
    Write-GuiLog "Listing installed applications from registry (safer than Win32_Product)..." -Color Cyan
    return {
        $ErrorActionPreference = 'Stop'
        $uninstallKeys = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        $installedApps = foreach ($keyPath in $uninstallKeys) {
            Get-ItemProperty -Path "$keyPath\*" -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and ($_.SystemComponent -ne 1 -or ($_.DisplayName -match "Visual C\+\+")) -and ($_.WindowsInstaller -ne 1 -or ($_.DisplayName -match "Visual C\+\+")) } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
            Sort-Object DisplayName -Unique
        }
        
        if ($installedApps) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Installed Applications (from Registry):"
            $installedApps | Format-Table -AutoSize | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil
        } else {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "No applications found or error retrieving list."
        }
    }
}

function Network-Diagnostics-ScriptBlock {
    Write-GuiLog "Performing network diagnostics..." -Color Cyan
    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Pinging external server (google.com)..."
        Test-Connection -ComputerName "google.com" -Count 4 | Format-Table -AutoSize | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil

        $localServers = @("10.60.70.11", "192.168.20.186")
        foreach ($server in $localServers) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Pinging local server $server..."
            Test-Connection -ComputerName $server -Count 4 -ErrorAction SilentlyContinue | Format-Table -AutoSize | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Network diagnostics completed."
    }
}

function Factory-Reset-Action {
    Write-GuiLog "This will reset the system to factory settings. ALL DATA WILL BE LOST!" -Color Red
    Write-GuiLog "WARNING: All personal files, apps, and settings will be removed. System will restart." -Color Red
    try {
        Write-GuiLog "Initiating factory reset..." -Color Cyan
        Start-Process -FilePath "systemreset.exe" -ArgumentList "-factoryreset" -Verb RunAs
        Write-GuiLog "Factory reset process started. Follow the on-screen prompts. The system will restart." -Color Green
    } catch {
        Write-GuiLog "Failed to initiate factory reset: $($_.Exception.Message)" -Color Red
    }
}

function Repair-Office-Action {
    Write-GuiLog "Attempting to repair Microsoft Office installation..." -Color Green
    $OfficeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $OfficeClickToRunPath) {
        try {
            Start-Process -FilePath $OfficeClickToRunPath -ArgumentList "scenario=Repair platform=x64 culture=en-us DisplayLevel=Full controlleaning=1" -Wait
            Write-GuiLog "Microsoft Office repair process initiated. Follow prompts if any." -Color Green
        } catch {
            Write-GuiLog "Error starting Office repair: $($_.Exception.Message)" -Color Red
        }
    } else {
        Write-GuiLog "Microsoft Office Click-to-Run client not found at $OfficeClickToRunPath." -Color Red
    }
}

function Check-OfficeUpdates-Action {
    Write-GuiLog "Checking for Microsoft Office updates..." -Color Green
    $OfficeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $OfficeClickToRunPath) {
        try {
            Start-Process -FilePath $OfficeClickToRunPath -ArgumentList "scenario=ApplyUpdates platform=x64 culture=en-us DisplayLevel=Full controlleaning=1" -Wait
            Write-GuiLog "Office update check initiated. Updates will be applied if available. Follow prompts." -Color Green
        } catch {
            Write-GuiLog "Error starting Office update check: $($_.Exception.Message)" -Color Red
        }
    } else {
        Write-GuiLog "Microsoft Office Click-to-Run client not found at $OfficeClickToRunPath." -Color Red
    }
}

function Update-Windows-ScriptBlock {
    Write-GuiLog "Checking for Windows updates using PSWindowsUpdate module..." -Color Cyan
    Write-GuiLog "This may install updates and automatically RESTART the computer." -Color Red
    return {
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "PSWindowsUpdate module not found. Attempting to install for CurrentUser..."
            try {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
                Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -AcceptLicense -Confirm:$false -ErrorAction Stop
                Microsoft.PowerShell.Host.WriteTranscriptUtil "PSWindowsUpdate module installed."
            } catch {
                 throw "Failed to install PSWindowsUpdate module: $($_.Exception.Message)"
            }
        }
        Import-Module PSWindowsUpdate -Force
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Searching for, downloading, and installing Windows updates (AcceptAll, AutoReboot)..."
        Get-WindowsUpdate -Install -AcceptAll -AutoReboot -Verbose:$false | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows update process completed. Check results above."
    }
}

function Clean-TempFiles-ScriptBlock {
    Write-GuiLog "Cleaning temporary files..." -Color Cyan
    return {
        $ErrorActionPreference = 'Continue'
        $tempPaths = @(
            "$env:TEMP\*", "C:\Windows\Temp\*", "$env:LOCALAPPDATA\Temp\*"
        )
        $cleanedCount = 0; $failedCount = 0
        foreach ($pathPattern in $tempPaths) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Processing path pattern: $pathPattern"
            $itemsToRemove = Get-ChildItem -Path $pathPattern -Recurse -Force -ErrorAction SilentlyContinue
            if ($itemsToRemove.Count -gt 0) {
                foreach ($item in $itemsToRemove) {
                    Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing: $($item.FullName)"
                    try { Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop }
                    catch { Microsoft.PowerShell.Host.WriteTranscriptUtil "WARN: Failed to clear '$($item.FullName)': $($_.Exception.Message)"; $failedCount++; Continue }
                    $cleanedCount++
                }
            } else {
                 Microsoft.PowerShell.Host.WriteTranscriptUtil "No items found for pattern: $pathPattern"
            }
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Temporary files cleanup completed. Items cleared: $cleanedCount, Failed: $failedCount."
    }
}

function Map-Printer-Action {
    param (
        [string]$PrinterConnectionName,
        [string]$PrinterFriendlyName
    )
    Write-GuiLog "Mapping printer: $PrinterFriendlyName ($PrinterConnectionName)..." -Color Green
    try {
        Add-Printer -ConnectionName $PrinterConnectionName -ErrorAction Stop
        Write-GuiLog "Printer '$PrinterFriendlyName' mapped successfully from $PrinterConnectionName." -Color Green
    } catch {
        Write-GuiLog "Failed to map printer '$PrinterFriendlyName' ($PrinterConnectionName): $($_.Exception.Message)" -Color Red
    }
}

function Install-AllPrinters-Action {
    Write-GuiLog "Installing all printers via VBS script..." -Color Green
    $vbsPaths = @(
        "\\server-mel\software\rp files\Printers.vbs",
        "\\server-syd\Scans\do not delete this folder\new pc files\Printers.vbs"
    )
    $vbsFound = $false
    foreach ($path in $vbsPaths) {
        if (Test-Path $path) {
            Write-GuiLog "Found Printers.vbs at $path" -Color Green
            try {
                Start-Process -FilePath "cscript.exe" -ArgumentList "//B //Nologo `"$path`"" -Wait
                Write-GuiLog "Printers installation script ($path) executed." -Color Green
                $vbsFound = $true; break
            } catch {
                Write-GuiLog "Failed to run Printers.vbs from ${path}: $($_.Exception.Message)" -Color Red
            }
        } else {
            Write-GuiLog "Printers.vbs not found or inaccessible at: $path" -Color Yellow
        }
    }
    if (-not $vbsFound) { Write-GuiLog "Printers.vbs script not found in any specified locations." -Color Red }
}

function Open-NewPCFiles-Action {
    Write-GuiLog "Opening New PC Files folder..." -Color Green
    $folderPath = '\\server-syd\Scans\do not delete this folder\new pc files'
    if(Test-Path $folderPath){
        try { Invoke-Item $folderPath; Write-GuiLog "Attempted to open folder: $folderPath" -Color Green }
        catch { Write-GuiLog "Failed to open folder $folderPath : $($_.Exception.Message)" -Color Red }
    } else { Write-GuiLog "Folder not found or inaccessible: $folderPath" -Color Red }
}

function Download-And-Open-Ninite-ScriptBlock {
    Write-GuiLog "Downloading Ninite installer..." -Color Green
    $niniteUrl = "https://ninite.com/.net4.8-.net4.8.1-7zip-chrome-vlc-zoom/ninite.exe"
    $outputPath = Join-Path $global:appsPath "ninite_rpi_custom.exe"
    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloading Ninite from $niniteUrl to $outputPath..."
        Invoke-WebRequest -Uri $niniteUrl -OutFile $outputPath -ErrorAction Stop
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloaded Ninite successfully. Opening installer..."
        Start-Process -FilePath $outputPath
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Ninite installer started. Please follow its prompts."
    }
}

function Download-MS-Teams-ScriptBlock {
    Write-GuiLog "Downloading MS Teams (New) installer and package..." -Color Green
    $teamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $teamsMsixUrl = "https://go.microsoft.com/fwlink/?linkid=2196106"
    $bootstrapperPath = Join-Path $global:appsPath "TeamsSetup_bootstrapper.exe"
    $msixPath = Join-Path $global:appsPath "MSTeams_x64.msix"

    return {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloading Teams Bootstrapper to $bootstrapperPath..."
        Invoke-WebRequest -Uri $teamsBootstrapperUrl -OutFile $bootstrapperPath -ErrorAction Stop
        
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloading Teams MSIX Package to $msixPath..."
        Invoke-WebRequest -Uri $teamsMsixUrl -OutFile $msixPath -ErrorAction Stop

        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloads complete. Waiting 5 seconds..."
        Start-Sleep -Seconds 5

        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running Teams Bootstrapper: $bootstrapperPath (may use downloaded MSIX: $msixPath)"
        Start-Process -FilePath $bootstrapperPath -Wait
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Microsoft Teams (New) installation initiated. Monitor for prompts."
    }
}

function Install-AdobeReader-ScriptBlock {
    Write-GuiLog "Installing Adobe Acrobat Reader DC (latest) using winget..." -Color Green
    Write-GuiLog "This will download and install Adobe Acrobat Reader. Ensure winget is installed and internet is active." -Color Yellow
    return {
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
            throw "winget command not found. Please install App Installer from Microsoft Store or ensure winget is in PATH."
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Installing Adobe Acrobat Reader DC via winget (ID: Adobe.Acrobat.Reader.DC)..."
        winget install --id Adobe.Acrobat.Reader.DC --exact --accept-source-agreements --accept-package-agreements --silent
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Adobe Acrobat Reader DC installation via winget initiated (silent)."
    }
}

function Remove-HPBloatware-ScriptBlock {
    Write-GuiLog "Removing HP bloatware and crapware..." -Color Green
    Write-GuiLog "This can take a while and might remove desired HP utilities if names overlap." -Color Yellow
    return {
        $ErrorActionPreference = 'Continue'
        
        $AppXPatternsToRemove = @(
            "*HPSupportAssistant*", "*HPJumpStarts*", "*HPPowerManager*", "*HPPrivacySettings*",
            "*HPSureShield*", "*HPQuickDrop*", "*HPWorkWell*", "*myHP*", "*HPDesktopSupportUtilities*",
            "*HPQuickTouch*", "*HPEasyClean*", "*HPPCHardwareDiagnosticsWindows*", "*HPSystemInformation*",
            "AD2F1837.*"
        )
        $ProgramNamePatternsToRemove = @(
            "HP Client Security Manager", "HP Connection Optimizer", "HP Documentation", "HP MAC Address Manager",
            "HP Notifications", "HP Security Update Service", "HP System Default Settings", "HP Sure Click",
            "HP Sure Run", "HP Sure Recover", "HP Sure Sense", "HP Wolf Security", "*HP Support Solutions Framework*"
        )

        Microsoft.PowerShell.Host.WriteTranscriptUtil "--- Removing AppX Provisioned Packages (HP related) ---"
        $ProvPackages = Get-AppxProvisionedPackage -Online
        foreach ($pattern in $AppXPatternsToRemove) {
            $matchingProv = $ProvPackages | Where-Object { $_.DisplayName -like $pattern -or $_.PackageName -like $pattern }
            foreach ($ProvPackage in $matchingProv) {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing provisioned package: $($ProvPackage.DisplayName) ($($ProvPackage.PackageName))"
                Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -AllUsers -ErrorAction SilentlyContinue
            }
        }

        Microsoft.PowerShell.Host.WriteTranscriptUtil "--- Removing AppX Packages (HP related) ---"
        $AppxUserPackages = Get-AppxPackage -AllUsers
        foreach ($pattern in $AppXPatternsToRemove) {
            $matchingAppx = $AppxUserPackages | Where-Object { $_.Name -like $pattern -or $_.PackageFullName -like $pattern }
            foreach ($AppxPackage in $matchingAppx) {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing AppX package: $($AppxPackage.Name) ($($AppxPackage.PackageFullName))"
                Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction SilentlyContinue
            }
        }
        
        Microsoft.PowerShell.Host.WriteTranscriptUtil "--- Uninstalling Programs (HP related) using Get-Package ---"
        $InstalledPrograms = Get-Package -ErrorAction SilentlyContinue
        foreach ($pattern in $ProgramNamePatternsToRemove) {
            $matchingPrograms = $InstalledPrograms | Where-Object { $_.Name -like $pattern }
            foreach ($Program in $matchingPrograms) {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Attempting to uninstall program: $($Program.Name) (Version: $($Program.Version))"
                try { $Program | Uninstall-Package -Force -ErrorAction Stop }
                catch { Microsoft.PowerShell.Host.WriteTranscriptUtil "WARN: Failed to uninstall $($Program.Name) via Uninstall-Package: $($_.Exception.Message)" }
            }
        }
        
        $MsiExecFallbacks = @('{0E2E04B0-9EDD-11EB-B38C-10604B96B11E}', '{4DA839F0-72CF-11EC-B247-3863BB3CB5A8}')
        foreach ($guid in $MsiExecFallbacks) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Attempting MSI uninstall for product code: $guid"
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/x `"$guid`" /qn /norestart" -Wait -ErrorAction SilentlyContinue
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "HP bloatware removal process completed. A restart might be beneficial."
    }
}


# --- GUI Construction ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "RPI Repair & Setup Tool v1.3" # Updated version
$mainForm.Size = New-Object System.Drawing.Size(900, 700)
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$mainForm.MinimumSize = New-Object System.Drawing.Size(750, 550)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel "Ready"
$statusStrip.Items.Add($statusLabel)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

$outputPanel = New-Object System.Windows.Forms.Panel
$outputPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$outputPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 200)
$outputPanel.Controls.Add($script:outputBox)

$mainForm.Controls.Clear()
$mainForm.Controls.AddRange(@($tabControl, $outputPanel, $statusStrip))
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
$outputPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

# Function to create a button (REVISED WITH .Tag)
function New-ToolButton {
    param(
        [string]$Text,
        $OnClickAction, # ScriptBlock for direct actions
        [System.Windows.Forms.Control]$ParentControl,
        [string]$JobScriptBlockFunctionName, # String name of function returning a ScriptBlock
        [string]$OperationNameForJob # String name for the job
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.AutoSize = $true
    $button.MinimumSize = New-Object System.Drawing.Size(250, 0)
    $button.Padding = New-Object System.Windows.Forms.Padding(10,5,10,5)
    $button.Margin = New-Object System.Windows.Forms.Padding(5)

    if ($JobScriptBlockFunctionName -and $OperationNameForJob) {
        $button.Tag = @{
            Type = "Job"
            JobFuncName = $JobScriptBlockFunctionName
            OpName = $OperationNameForJob
        }
        $button.Add_Click({
            param($sender, $eventArgs) # Standard event signature
            $clickedButton = $sender -as [System.Windows.Forms.Button]
            $buttonData = $clickedButton.Tag

            $statusLabel.Text = "Starting job: $($buttonData.OpName)..."
            Write-GuiLog "Preparing job '$($buttonData.OpName)'. ScriptBlock function: '$($buttonData.JobFuncName)'." -Color DarkCyan

            if (-not $buttonData.JobFuncName) {
                Write-GuiLog "CRITICAL ERROR (Button: $($clickedButton.Text)): JobFuncName in Tag is null or empty for operation '$($buttonData.OpName)'." -Color Red
                $statusLabel.Text = "Error preparing job: $($buttonData.OpName)"
                return
            }

            $scriptBlockFromFunc = $null
            try {
                $scriptBlockFromFunc = Invoke-Expression $buttonData.JobFuncName -ErrorAction Stop
            } catch {
                Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Failed to invoke '$($buttonData.JobFuncName)' for operation '$($buttonData.OpName)'. Error: $($_.Exception.Message)" -Color Red
                $statusLabel.Text = "Error invoking script function for: $($buttonData.OpName)"
                return
            }
            
            if (-not $scriptBlockFromFunc) {
                Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Function '$($buttonData.JobFuncName)' did NOT return a scriptblock (returned null) for operation '$($buttonData.OpName)'." -Color Red
                $statusLabel.Text = "Error: Script function returned null for: $($buttonData.OpName)"
                return
            }
            if ($scriptBlockFromFunc -isnot [scriptblock]) {
                 Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Function '$($buttonData.JobFuncName)' returned a value of type '$($scriptBlockFromFunc.GetType().FullName)' instead of [scriptblock] for operation '$($buttonData.OpName)'." -Color Red
                 $statusLabel.Text = "Error: Script function returned wrong type for: $($buttonData.OpName)"
                 return
            }
            Start-LongRunningJob -ScriptBlock $scriptBlockFromFunc -OperationName $buttonData.OpName -ButtonToDisable $clickedButton
        })

    } elseif ($OnClickAction -is [scriptblock]) {
        $button.Tag = @{
            Type = "Direct"
            Action = $OnClickAction # Store the actual scriptblock
        }
        $button.Add_Click({
            param($sender, $eventArgs)
            $clickedButton = $sender -as [System.Windows.Forms.Button]
            $buttonData = $clickedButton.Tag

            $statusLabel.Text = "Executing: $($clickedButton.Text)..."
            try {
                if ($buttonData.Action -is [scriptblock]) {
                    Invoke-Command -ScriptBlock $buttonData.Action
                    Write-GuiLog "Direct action '$($clickedButton.Text)' completed." -Color Green
                } else {
                    Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Action in Tag is not a scriptblock. Type: $($buttonData.Action.GetType().FullName)" -Color Red
                }
            } catch {
                Write-GuiLog "Error during direct action '$($clickedButton.Text)': $($_.Exception.Message)" -Color Red
                if ($buttonData.Action -is [scriptblock]) {
                    Write-GuiLog "Failing scriptblock for direct action was: $($buttonData.Action.ToString())" -Color DarkRed
                }
            }
            $statusLabel.Text = "Finished: $($clickedButton.Text). Check log."
        })
    } else {
        $button.Text = "$Text (Misconfigured)"
        $button.Enabled = $false
        Write-GuiLog "Button '$Text' is misconfigured: No valid job function name/operation name or direct action scriptblock provided." -Color Magenta
    }
    $ParentControl.Controls.Add($button)
    return $button
}


function New-ButtonFlowPanel {
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $panel.AutoScroll = $true
    $panel.WrapContents = $false
    $panel.Padding = New-Object System.Windows.Forms.Padding(15)
    return $panel
}

# == New PC Setup Tab ==
$tabNewPC = New-Object System.Windows.Forms.TabPage; $tabNewPC.Text = "New PC Setup"
$panelNewPC = New-ButtonFlowPanel; $tabNewPC.Controls.Add($panelNewPC)
New-ToolButton "1: Open New PC Files Folder" {Open-NewPCFiles-Action} $panelNewPC $null $null
New-ToolButton "2: Download & Run Ninite" $null $panelNewPC "Download-And-Open-Ninite-ScriptBlock" "Ninite Download & Run"
New-ToolButton "3: Download MS Teams (New)" $null $panelNewPC "Download-MS-Teams-ScriptBlock" "MS Teams Download"
New-ToolButton "4: Change PC Name (Restarts PC)" {Change-PCName-Action} $panelNewPC $null $null

$groupJoinDomain = New-Object System.Windows.Forms.GroupBox; $groupJoinDomain.Text = "5: Join RPI Domain (Restarts PC)"; $groupJoinDomain.AutoSize = $true; $groupJoinDomain.Padding = New-Object System.Windows.Forms.Padding(10)
$flowJoinDomain = New-Object System.Windows.Forms.FlowLayoutPanel; $flowJoinDomain.Dock = [System.Windows.Forms.DockStyle]::Fill; $flowJoinDomain.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown; $flowJoinDomain.AutoSize = $true
$lblDomain = New-Object System.Windows.Forms.Label; $lblDomain.Text = "Domain Name (e.g., rootprojects.local):"; $lblDomain.AutoSize=$true
$txtDomainName = New-Object System.Windows.Forms.TextBox; $txtDomainName.Width = 250; $txtDomainName.Text = "rootprojects.local"
$lblOU = New-Object System.Windows.Forms.Label; $lblOU.Text = "OU Path (e.g., OU=Computers,DC=rootprojects,DC=local):"; $lblOU.AutoSize=$true
$txtOUPath = New-Object System.Windows.Forms.TextBox; $txtOUPath.Width = 350; $txtOUPath.Text = "OU=Computers,DC=rootprojects,DC=local"
$btnJoinDomain = New-Object System.Windows.Forms.Button; $btnJoinDomain.Text = "Join Domain"; $btnJoinDomain.Padding = New-Object System.Windows.Forms.Padding(5)
$btnJoinDomain.Add_Click({
    $domain = $txtDomainName.Text.Trim(); $ou = $txtOUPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($domain)) { [System.Windows.Forms.MessageBox]::Show("Domain Name required.", "Error", "OK", "Error"); return }
    try {
        $credential = Get-Credential -Message "Enter credentials for '$domain' domain join"
        if ($credential) { Join-Domain-Action -domainName $domain -ouPath $ou -credential $credential }
        else { Write-GuiLog "Domain join cancelled - no credentials." -Color Yellow }
    } catch { Write-GuiLog "Domain join failed or cancelled during credential prompt: $($_.Exception.Message)" -Color Yellow }
})
$flowJoinDomain.Controls.AddRange(@($lblDomain, $txtDomainName, $lblOU, $txtOUPath, $btnJoinDomain))
$groupJoinDomain.Controls.Add($flowJoinDomain); $panelNewPC.Controls.Add($groupJoinDomain)

New-ToolButton "6: Update Windows (May Restart)" $null $panelNewPC "Update-Windows-ScriptBlock" "Windows Update"
New-ToolButton "7: Install Adobe Acrobat Reader" $null $panelNewPC "Install-AdobeReader-ScriptBlock" "Adobe Reader Install"
New-ToolButton "8: Remove HP Bloatware" $null $panelNewPC "Remove-HPBloatware-ScriptBlock" "HP Bloatware Removal"
$tabControl.Controls.Add($tabNewPC)

# == Windows Repairs Tab ==
$tabWindows = New-Object System.Windows.Forms.TabPage; $tabWindows.Text = "Windows Repairs"
$panelWindows = New-ButtonFlowPanel; $tabWindows.Controls.Add($panelWindows)
New-ToolButton "1: DISM RestoreHealth (May Restart)" $null $panelWindows "Repair-Windows-ScriptBlock" "DISM RestoreHealth"
New-ToolButton "2: System File Checker (SFC)" $null $panelWindows "Repair-SystemFiles-ScriptBlock" "SFC Scan"
New-ToolButton "3: Check Disk (CHKDSK C:) (Restarts)" {Repair-Disk-Action} $panelWindows $null $null
New-ToolButton "4: Windows Update Troubleshooter" {Run-WindowsUpdateTroubleshooter-Action} $panelWindows $null $null
New-ToolButton "5: DISM Check & Repair (May Restart)" $null $panelWindows "Check-And-Repair-DISM-ScriptBlock" "DISM Check & Repair"
New-ToolButton "6: Network Reset (May Restart)" $null $panelWindows "Reset-Network-ScriptBlock" "Network Reset"
New-ToolButton "7: Windows Memory Diagnostic (Restarts)" {Run-MemoryDiagnostic-Action} $panelWindows $null $null
New-ToolButton "8: Windows Startup Repair (Restarts)" {Run-StartupRepair-Action} $panelWindows $null $null
New-ToolButton "9: Windows Defender Full Scan" $null $panelWindows "Run-WindowsDefenderScan-ScriptBlock" "Defender Full Scan"
New-ToolButton "10: Reset Windows Update Components" $null $panelWindows "Reset-WindowsUpdateComponents-ScriptBlock" "Reset WU Components"
New-ToolButton "11: List Installed Apps" $null $panelWindows "List-InstalledApps-ScriptBlock" "List Installed Apps"
New-ToolButton "12: Network Diagnostics" $null $panelWindows "Network-Diagnostics-ScriptBlock" "Network Diagnostics"
New-ToolButton "13: Factory Reset (Data Loss & Restarts!)" {Factory-Reset-Action} $panelWindows $null $null
$tabControl.Controls.Add($tabWindows)

# == Office Repairs Tab ==
$tabOffice = New-Object System.Windows.Forms.TabPage; $tabOffice.Text = "Office Repairs"
$panelOffice = New-ButtonFlowPanel; $tabOffice.Controls.Add($panelOffice)
New-ToolButton "1: Repair Microsoft Office" {Repair-Office-Action} $panelOffice $null $null
New-ToolButton "2: Check for Microsoft Office Updates" {Check-OfficeUpdates-Action} $panelOffice $null $null
$tabControl.Controls.Add($tabOffice)

# == User Tasks Tab ==
$tabUserTasks = New-Object System.Windows.Forms.TabPage; $tabUserTasks.Text = "User Tasks"
$panelUserTasks = New-ButtonFlowPanel; $tabUserTasks.Controls.Add($panelUserTasks)
New-ToolButton "1: Clean Temp Files" $null $panelUserTasks "Clean-TempFiles-ScriptBlock" "Clean Temp Files"

$groupPrinters = New-Object System.Windows.Forms.GroupBox; $groupPrinters.Text = "2: Printer Mapping"; $groupPrinters.AutoSize = $true; $groupPrinters.Padding = New-Object System.Windows.Forms.Padding(10)
$flowPrinters = New-Object System.Windows.Forms.FlowLayoutPanel; $flowPrinters.Dock = [System.Windows.Forms.DockStyle]::Fill; $flowPrinters.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown; $flowPrinters.AutoSize = $true

$printersToMap = @{
    "Sydney Printer (192.168.23.10)" = "\\192.168.23.10"
    "Melbourne Printer (192.168.33.63)" = "\\192.168.33.63"
    "Melb Airport Printer (192.168.43.250)" = "\\192.168.43.250"
    "Townsville Printer (192.168.100.240)" = "\\192.168.100.240"
    "Brisbane Printer (192.168.20.242)" = "\\192.168.20.242"
    "Mackay Printer (192.168.90.240)" = "\\192.168.90.240"
}
foreach ($entry in $printersToMap.GetEnumerator()) {
    $actionScriptBlock = [scriptblock]::Create("Map-Printer-Action -PrinterConnectionName '$($entry.Value)' -PrinterFriendlyName '$($entry.Key)'")
    New-ToolButton "Map $($entry.Key)" $actionScriptBlock $flowPrinters $null $null
}
New-ToolButton "Install All Printers (via VBS)" {Install-AllPrinters-Action} $flowPrinters $null $null
$groupPrinters.Controls.Add($flowPrinters); $panelUserTasks.Controls.Add($groupPrinters)

New-ToolButton "3: Clear Teams Cache & Restart Teams" $null $panelUserTasks "Clear-TeamsCache-ScriptBlock" "Clear Teams Cache"
New-ToolButton "BONUS: Start Microsoft Teams" {Start-Teams-Action} $panelUserTasks $null $null
$tabControl.Controls.Add($tabUserTasks)


# --- Form Load and Show ---
$mainForm.Add_Shown({
    Write-GuiLog "RPI Repair & Setup Tool GUI Initialized." -Color Blue
    $statusLabel.Text = "Ready."
})
$mainForm.Add_FormClosing({
    Write-GuiLog "Exiting RPI Repair Tool..." -Color Blue
    Get-EventSubscriber -SourceIdentifier JobEvent_* -ErrorAction SilentlyContinue | ForEach-Object { Unregister-Event -SubscriptionId $_.SubscriptionId }
    Get-Job | Where-Object {$_.Name -like "*"} | Remove-Job -Force -ErrorAction SilentlyContinue
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$mainForm.ShowDialog()
