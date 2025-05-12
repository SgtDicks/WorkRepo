#Requires -Version 5.1
#Requires -RunAsAdministrator

# --- INITIAL SETUP (Console Recap Disabled) ---
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser -ErrorAction SilentlyContinue
    if ($currentPolicy -ne "RemoteSigned" -and $currentPolicy -ne "Unrestricted") {
        # Write-Host "Current Execution Policy for CurrentUser is $currentPolicy. Setting to RemoteSigned." # CONSOLE RECAP DISABLED
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
        # Write-Host "Execution Policy for CurrentUser set to RemoteSigned." -ForegroundColor Green # CONSOLE RECAP DISABLED
    } else {
        # Write-Host "Execution Policy for CurrentUser is already sufficient ($currentPolicy)." -ForegroundColor Yellow # CONSOLE RECAP DISABLED
    }
} catch {
    # Write-Warning "Failed to set Execution Policy. $_. Some script features might not work." # CONSOLE RECAP DISABLED
}

$global:appsPath = "C:\apps" 
if (-not (Test-Path -Path $global:appsPath)) {
    try {
        New-Item -ItemType Directory -Path $global:appsPath -ErrorAction Stop | Out-Null
        # Write-Host "Created directory: $($global:appsPath)" -ForegroundColor Green # CONSOLE RECAP DISABLED
    } catch {
        # Write-Warning "Failed to create directory $($global:appsPath). $_. Some downloads may fail." # CONSOLE RECAP DISABLED
    }
} else {
    # Write-Host "Directory already exists: $($global:appsPath)" -ForegroundColor Yellow # CONSOLE RECAP DISABLED
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

# --- FULL OPERATIONAL FUNCTION DEFINITIONS (WITH $using: FIXES) ---

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
        $evtJob = $Sender; $jobState = $evtJob.JobStateInfo.State; $sourceId = $EventArgs.SourceIdentifier
        $script:outputBox.Invoke([Action]{
            Write-GuiLog "Job '$($evtJob.Name)' state: $jobState" -Color Gray
            if ($jobState -in ('Completed', 'Failed', 'Stopped')) {
                $jobErrors = $evtJob.ChildJobs[0].Error
                if ($jobErrors.Count -gt 0) { Write-GuiLog "Errors from job '$($evtJob.Name)':" -Color Red; $jobErrors | ForEach-Object { Write-GuiLog $_.ToString() -Color Red -NoTimestamp } }
                $output = Receive-Job -Job $evtJob -Keep
                if ($output) { Write-GuiLog "Output from job '$($evtJob.Name)':" -Color DarkGray; $output | ForEach-Object { Write-GuiLog $_.ToString() -Color DarkGray -NoTimestamp } }
                if ($jobState -eq 'Completed') { Write-GuiLog "Job '$($evtJob.Name)' completed successfully." -Color Green }
                else { Write-GuiLog "Job '$($evtJob.Name)' finished with state: $jobState. Reason: $($evtJob.JobStateInfo.Reason)" -Color Red }
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
    } catch { Write-GuiLog "Error changing PC name: $($_.Exception.Message)" -Color Red }
}

function Join-Domain-Action {
    param ([string]$domainName, [string]$ouPath, [System.Management.Automation.PSCredential]$credential)
    Write-GuiLog "Joining the computer to the domain $domainName..." -Color Cyan
    if ($ouPath) { Write-GuiLog "Target OU Path: $ouPath" -Color Cyan } else { Write-GuiLog "No OU Path specified, using default computer container." -Color Yellow }
    Write-GuiLog "Computer will be joined to domain and RESTART. Ensure all work is saved." -Color Red
    $commandParams = @{ DomainName = $domainName; Credential = $credential; Force = $true; Restart = $true }
    if (-not [string]::IsNullOrWhiteSpace($ouPath)) { $commandParams.OUPath = $ouPath }
    try { Add-Computer @commandParams -ErrorAction Stop; Write-GuiLog "Join domain command issued. System should restart shortly." -Color Green }
    catch {
        if ($_.Exception -is [System.InvalidOperationException] -and ($_.Exception.Message -like "*already in that domain*" -or $_.Exception.Message -like "*already a member of domain*")) { Write-GuiLog "Computer is already a member of the domain '$domainName'." -Color Yellow }
        else { Write-GuiLog "Error joining domain '$domainName': $($_.Exception.Message)" -Color Red }
    }
}

function Repair-Windows-ScriptBlock {
    Write-GuiLog "Running DISM command to restore health..." -Color Green
    Write-GuiLog "This process can take 15-30 minutes. A restart may be required." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting DISM /Online /Cleanup-Image /RestoreHealth..."
        dism /online /cleanup-image /restorehealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows repair (DISM RestoreHealth) completed."
    }
    return $scriptBlock
}

function Repair-SystemFiles-ScriptBlock {
    Write-GuiLog "Running System File Checker (SFC)..." -Color Green
    Write-GuiLog "This scan can take up to 20 minutes. No restart is required unless issues are found." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting SFC /scannow..."
        sfc /scannow
        Microsoft.PowerShell.Host.WriteTranscriptUtil "System File Checker (SFC) completed."
    }
    return $scriptBlock
}

function Repair-Disk-Action {
    Write-GuiLog "Running Check Disk (CHKDSK) on C:..." -Color Green
    Write-GuiLog "This operation could take a few hours and WILL RESTART the computer. Ensure all work is saved." -Color Red
    try { chkdsk C: /F /R /X; Write-GuiLog "CHKDSK scheduled for the next restart." -Color Green }
    catch { Write-GuiLog "Error scheduling CHKDSK: $($_.Exception.Message)" -Color Red }
}

function Run-WindowsUpdateTroubleshooter-Action {
    Write-GuiLog "Running Windows Update Troubleshooter..." -Color Green
    Write-GuiLog "This operation might take 5-10 minutes. An interactive window will open." -Color Yellow
    try { Start-Process -FilePath "msdt.exe" -ArgumentList "/id WindowsUpdateDiagnostic"; Write-GuiLog "Windows Update Troubleshooter started." -Color Green }
    catch { Write-GuiLog "Failed to start Windows Update Troubleshooter: $($_.Exception.Message)" -Color Red }
}

function Check-And-Repair-DISM-ScriptBlock {
    Write-GuiLog "Running DISM Check and Repair..." -Color Green
    Write-GuiLog "This process might take 30-60 minutes. A restart may be required." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running DISM /Online /Cleanup-Image /CheckHealth..."
        DISM /Online /Cleanup-Image /CheckHealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running DISM /Online /Cleanup-Image /ScanHealth..."
        DISM /Online /Cleanup-Image /ScanHealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running DISM /Online /Cleanup-Image /RestoreHealth..."
        DISM /Online /Cleanup-Image /RestoreHealth
        Microsoft.PowerShell.Host.WriteTranscriptUtil "DISM Check and Repair completed."
    }
    return $scriptBlock
}

function Reset-Network-ScriptBlock {
    Write-GuiLog "Resetting Network Adapters..." -Color Green
    Write-GuiLog "This will cause a temporary network outage and may require a restart." -Color Red
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        $Commands = @({ netsh winsock reset }, { netsh int ip reset }, { ipconfig /release }, { ipconfig /renew }, { ipconfig /flushdns })
        foreach($cmd in $Commands){ Microsoft.PowerShell.Host.WriteTranscriptUtil "Executing: $($cmd.ToString())"; Invoke-Command -ScriptBlock $cmd }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Network reset commands executed. A restart may be required."
    }
    return $scriptBlock
}

function Run-MemoryDiagnostic-Action {
    Write-GuiLog "Scheduling Windows Memory Diagnostic..." -Color Green
    Write-GuiLog "This test WILL RESTART your computer. Ensure all work is saved." -Color Red
    try { Start-Process -FilePath "mdsched.exe" -ArgumentList "/f" -Verb RunAs; Write-GuiLog "Windows Memory Diagnostic scheduled." -Color Green }
    catch { Write-GuiLog "Failed to schedule Windows Memory Diagnostic: $($_.Exception.Message)" -Color Red }
}

function Run-StartupRepair-Action {
    Write-GuiLog "Initiating Startup Repair..." -Color Green
    Write-GuiLog "This process WILL RESTART your computer." -Color Red
    try {
        Write-GuiLog "Configuring boot to recovery..." -Color Cyan; Start-Process -FilePath "reagentc.exe" -ArgumentList "/boottore" -Verb RunAs -Wait
        Write-GuiLog "Restarting NOW..." -Color Cyan; Shutdown.exe /r /t 0 /f
    } catch { Write-GuiLog "Failed to initiate Startup Repair: $($_.Exception.Message)" -Color Red }
}

function Run-WindowsDefenderScan-ScriptBlock {
    Write-GuiLog "Running Windows Defender Full Scan..." -Color Green
    Write-GuiLog "This scan can take several hours." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting Windows Defender Full Scan..."
        Start-MpScan -ScanType FullScan -ErrorAction Stop
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows Defender Full Scan completed."
    }
    return $scriptBlock
}

function Reset-WindowsUpdateComponents-ScriptBlock {
    Write-GuiLog "Resetting Windows Update components..." -Color Green
    Write-GuiLog "This may take 10-20 minutes." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        $servicesToManage = @("wuauserv", "cryptSvc", "bits", "msiserver")
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Stopping services: $($servicesToManage -join ', ')"
        Stop-Service -Name $servicesToManage -Force -ErrorAction SilentlyContinue
        $pathsToRename = @{ "C:\Windows\SoftwareDistribution" = "SoftwareDistribution.old"; "C:\Windows\System32\catroot2" = "catroot2.old" }
        foreach ($entry in $pathsToRename.GetEnumerator()) {
            $oldPath = $entry.Key; $newDirName = $entry.Value; $parentDir = Split-Path $oldPath; $newFullPath = Join-Path $parentDir $newDirName
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Renaming $oldPath to $newFullPath"
            if (Test-Path $newFullPath) { Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing $newFullPath..."; Remove-Item -Path $newFullPath -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path $oldPath) { Rename-Item -Path $oldPath -NewName $newDirName -Force -ErrorAction Stop; Microsoft.PowerShell.Host.WriteTranscriptUtil "Renamed $oldPath" }
            else { Microsoft.PowerShell.Host.WriteTranscriptUtil "$oldPath not found." }
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Starting services: $($servicesToManage -join ', ')"
        Start-Service -Name $servicesToManage -ErrorAction SilentlyContinue
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows Update components reset completed."
    }
    return $scriptBlock
}

function Start-Teams-Action {
    Write-GuiLog "Trying to find and launch Microsoft Teams..." -Color Cyan
    $possiblePaths = @("$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe", (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\ms-teams.exe"), "$env:LOCALAPPDATA\Programs\Teams\current\Teams.exe", "$env:PROGRAMFILES\Teams Installer\Teams.exe", "$env:PROGRAMFILES(X86)\Teams Installer\Teams.exe")
    $teamsPath = $null
    foreach ($path in $possiblePaths) {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
        if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) { $teamsPath = $resolvedPath; Write-GuiLog "Teams found: $teamsPath" -Color Green; break }
        else { Write-GuiLog "Teams not at: $resolvedPath" -Color DarkGray }
    }
    if ($teamsPath) { try { Start-Process -FilePath $teamsPath; Write-GuiLog "Teams launched." -Color Green } catch { Write-GuiLog "Failed to start Teams from ${teamsPath}: $($_.Exception.Message)" -Color Red } }
    else { Write-GuiLog "Teams executable not found." -Color Red }
}

function Clear-TeamsCache-ScriptBlock {
    Write-GuiLog "Attempting to clear Microsoft Teams cache..." -Color Cyan
    $scriptBlock = {
        $ErrorActionPreference = 'Continue'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Closing Teams processes..."
        Get-Process -Name "ms-teams", "Teams" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3
        $cacheLocations = @("$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe", "$env:LOCALAPPDATA\Packages\MicrosoftTeams_8wekyb3d8bbwe", "$env:LOCALAPPDATA\Microsoft\Teams", "$env:APPDATA\Microsoft\Teams")
        foreach ($location in $cacheLocations) {
            $resolvedLocation = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($location)
            if (Test-Path -LiteralPath $resolvedLocation) {
                Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing: $resolvedLocation"
                try { Remove-Item -Path $resolvedLocation -Recurse -Force -ErrorAction Stop; Microsoft.PowerShell.Host.WriteTranscriptUtil "Removed: $resolvedLocation" }
                catch { Microsoft.PowerShell.Host.WriteTranscriptUtil "WARN: Failed to remove $resolvedLocation. $($_.Exception.Message)" }
            } else { Microsoft.PowerShell.Host.WriteTranscriptUtil "Path not found: $resolvedLocation" }
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Teams cache clear done. Restarting Teams..."
        $possiblePathsToStart = @("$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe", (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\ms-teams.exe"))
        $teamsExeToStart = $null
        foreach ($p in $possiblePathsToStart) { $rp = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p); if (Test-Path -LiteralPath $rp -PathType Leaf) { $teamsExeToStart = $rp; break } }
        if ($teamsExeToStart) { Start-Process -FilePath $teamsExeToStart; Microsoft.PowerShell.Host.WriteTranscriptUtil "Started Teams from $teamsExeToStart." }
        else { Microsoft.PowerShell.Host.WriteTranscriptUtil "Could not auto-restart Teams." }
    }
    return $scriptBlock
}

function List-InstalledApps-ScriptBlock {
    Write-GuiLog "Listing installed applications from registry..." -Color Cyan
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        $uninstallKeys = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall")
        $installedApps = foreach ($keyPath in $uninstallKeys) { Get-ItemProperty -Path "$keyPath\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and ($_.SystemComponent -ne 1 -or ($_.DisplayName -match "Visual C\+\+")) -and ($_.WindowsInstaller -ne 1 -or ($_.DisplayName -match "Visual C\+\+")) } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object DisplayName -Unique }
        if ($installedApps) { Microsoft.PowerShell.Host.WriteTranscriptUtil "Installed Apps:"; $installedApps | Format-Table -AutoSize | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil }
        else { Microsoft.PowerShell.Host.WriteTranscriptUtil "No apps found or error." }
    }
    return $scriptBlock
}

function Network-Diagnostics-ScriptBlock {
    Write-GuiLog "Performing network diagnostics..." -Color Cyan
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Pinging google.com..."
        Test-Connection -ComputerName "google.com" -Count 4 | Format-Table -AutoSize | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil
        $localServers = @("10.60.70.11", "192.168.20.186")
        foreach ($server in $localServers) { Microsoft.PowerShell.Host.WriteTranscriptUtil "Pinging $server..."; Test-Connection -ComputerName $server -Count 4 -ErrorAction SilentlyContinue | Format-Table -AutoSize | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Network diagnostics done."
    }
    return $scriptBlock
}

function Factory-Reset-Action {
    Write-GuiLog "This will reset the system. ALL DATA WILL BE LOST!" -Color Red
    Write-GuiLog "WARNING: Files, apps, settings removed. System RESTART." -Color Red
    try { Write-GuiLog "Initiating factory reset..." -Color Cyan; Start-Process -FilePath "systemreset.exe" -ArgumentList "-factoryreset" -Verb RunAs; Write-GuiLog "Factory reset started. Follow prompts." -Color Green }
    catch { Write-GuiLog "Failed to initiate factory reset: $($_.Exception.Message)" -Color Red }
}

function Repair-Office-Action {
    Write-GuiLog "Attempting to repair Microsoft Office..." -Color Green
    $OfficeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $OfficeClickToRunPath) { try { Start-Process -FilePath $OfficeClickToRunPath -ArgumentList "scenario=Repair platform=x64 culture=en-us DisplayLevel=Full controlleaning=1" -Wait; Write-GuiLog "Office repair initiated." -Color Green } catch { Write-GuiLog "Error starting Office repair: $($_.Exception.Message)" -Color Red } }
    else { Write-GuiLog "Office Click-to-Run client not found." -Color Red }
}

function Check-OfficeUpdates-Action {
    Write-GuiLog "Checking for Microsoft Office updates..." -Color Green
    $OfficeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $OfficeClickToRunPath) { try { Start-Process -FilePath $OfficeClickToRunPath -ArgumentList "scenario=ApplyUpdates platform=x64 culture=en-us DisplayLevel=Full controlleaning=1" -Wait; Write-GuiLog "Office update check initiated." -Color Green } catch { Write-GuiLog "Error starting Office update check: $($_.Exception.Message)" -Color Red } }
    else { Write-GuiLog "Office Click-to-Run client not found." -Color Red }
}

function Update-Windows-ScriptBlock {
    Write-GuiLog "Checking for Windows updates (PSWindowsUpdate)..." -Color Cyan
    Write-GuiLog "This may install updates and RESTART computer." -Color Red
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "PSWindowsUpdate module not found. Installing..."
            try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop; Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber -AcceptLicense -Confirm:$false -ErrorAction Stop; Microsoft.PowerShell.Host.WriteTranscriptUtil "PSWindowsUpdate installed." }
            catch { throw "Failed to install PSWindowsUpdate: $($_.Exception.Message)" }
        }
        Import-Module PSWindowsUpdate -Force
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Installing Windows updates (AcceptAll, AutoReboot)..."
        Get-WindowsUpdate -Install -AcceptAll -AutoReboot -Verbose:$false | Out-String | Microsoft.PowerShell.Host.WriteTranscriptUtil
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Windows update process done."
    }
    return $scriptBlock
}

function Clean-TempFiles-ScriptBlock {
    Write-GuiLog "Cleaning temporary files..." -Color Cyan
    $scriptBlock = {
        $ErrorActionPreference = 'Continue'
        $tempPaths = @("$env:TEMP\*", "C:\Windows\Temp\*", "$env:LOCALAPPDATA\Temp\*")
        $cleanedCount = 0; $failedCount = 0
        foreach ($pathPattern in $tempPaths) {
            Microsoft.PowerShell.Host.WriteTranscriptUtil "Processing: $pathPattern"
            $itemsToRemove = Get-ChildItem -Path $pathPattern -Recurse -Force -ErrorAction SilentlyContinue
            if ($itemsToRemove.Count -gt 0) { foreach ($item in $itemsToRemove) { Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing: $($item.FullName)"; try { Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop } catch { Microsoft.PowerShell.Host.WriteTranscriptUtil "WARN: Failed: '$($item.FullName)': $($_.Exception.Message)"; $failedCount++; Continue }; $cleanedCount++ } }
            else { Microsoft.PowerShell.Host.WriteTranscriptUtil "No items for: $pathPattern" }
        }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Temp files cleanup done. Cleared: $cleanedCount, Failed: $failedCount."
    }
    return $scriptBlock
}

function Map-Printer-Action {
    param ([string]$PrinterConnectionName, [string]$PrinterFriendlyName)
    Write-GuiLog "Mapping printer: $PrinterFriendlyName ($PrinterConnectionName)..." -Color Green
    try { Add-Printer -ConnectionName $PrinterConnectionName -ErrorAction Stop; Write-GuiLog "Printer '$PrinterFriendlyName' mapped from $PrinterConnectionName." -Color Green }
    catch { Write-GuiLog "Failed to map printer '$PrinterFriendlyName' ($PrinterConnectionName): $($_.Exception.Message)" -Color Red }
}

function Install-AllPrinters-Action {
    Write-GuiLog "Installing all printers via VBS script..." -Color Green
    $vbsPaths = @("\\server-mel\software\rp files\Printers.vbs", "\\server-syd\Scans\do not delete this folder\new pc files\Printers.vbs")
    $vbsFound = $false
    foreach ($path in $vbsPaths) {
        if (Test-Path $path) { Write-GuiLog "Found Printers.vbs: $path" -Color Green; try { Start-Process -FilePath "cscript.exe" -ArgumentList "//B //Nologo `"$path`"" -Wait; Write-GuiLog "Printers script ($path) executed." -Color Green; $vbsFound = $true; break } catch { Write-GuiLog "Failed to run Printers.vbs from ${path}: $($_.Exception.Message)" -Color Red } }
        else { Write-GuiLog "Printers.vbs not at: $path" -Color Yellow }
    }
    if (-not $vbsFound) { Write-GuiLog "Printers.vbs script not found." -Color Red }
}

function Open-NewPCFiles-Action {
    Write-GuiLog "Opening New PC Files folder..." -Color Green
    $folderPath = '\\server-syd\Scans\do not delete this folder\new pc files'
    if(Test-Path $folderPath){ try { Invoke-Item $folderPath; Write-GuiLog "Opened: $folderPath" -Color Green } catch { Write-GuiLog "Failed to open $folderPath : $($_.Exception.Message)" -Color Red } }
    else { Write-GuiLog "Folder not found: $folderPath" -Color Red }
}

function Download-And-Open-Ninite-ScriptBlock {
    Write-GuiLog "Downloading Ninite installer..." -Color Green
    $niniteUrl = "https://ninite.com/.net4.8-.net4.8.1-7zip-chrome-vlc-zoom/ninite.exe"
    $localAppsPath = $global:appsPath # Capture for $using:
    $outputPath = Join-Path $localAppsPath "ninite_rpi_custom.exe"
    
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloading Ninite from $($using:niniteUrl) to $($using:outputPath)..."
        Invoke-WebRequest -Uri $using:niniteUrl -OutFile $using:outputPath -ErrorAction Stop
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloaded Ninite. Opening installer..."
        Start-Process -FilePath $using:outputPath
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Ninite installer started."
    }
    return $scriptBlock
}

function Download-MS-Teams-ScriptBlock {
    Write-GuiLog "Downloading MS Teams (New) installer and package..." -Color Green
    $teamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $teamsMsixUrl = "https://go.microsoft.com/fwlink/?linkid=2196106"
    $localAppsPath = $global:appsPath # Capture for $using:
    $bootstrapperPath = Join-Path $localAppsPath "TeamsSetup_bootstrapper.exe"
    $msixPath = Join-Path $localAppsPath "MSTeams_x64.msix"

    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloading Teams Bootstrapper to $($using:bootstrapperPath)..."
        Invoke-WebRequest -Uri $using:teamsBootstrapperUrl -OutFile $using:bootstrapperPath -ErrorAction Stop
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloading Teams MSIX Package to $($using:msixPath)..."
        Invoke-WebRequest -Uri $using:teamsMsixUrl -OutFile $using:msixPath -ErrorAction Stop
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Downloads complete. Waiting 5s..."
        Start-Sleep -Seconds 5
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Running Teams Bootstrapper: $($using:bootstrapperPath)..."
        Start-Process -FilePath $using:bootstrapperPath -Wait
        Microsoft.PowerShell.Host.WriteTranscriptUtil "MS Teams (New) installation initiated."
    }
    return $scriptBlock
}

function Install-AdobeReader-ScriptBlock {
    Write-GuiLog "Installing Adobe Acrobat Reader DC (winget)..." -Color Green
    Write-GuiLog "Ensuring winget and internet." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { throw "winget not found." }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Installing Adobe Acrobat Reader DC (winget)..."
        winget install --id Adobe.Acrobat.Reader.DC --exact --accept-source-agreements --accept-package-agreements --silent
        Microsoft.PowerShell.Host.WriteTranscriptUtil "Adobe Reader install (winget) initiated."
    }
    return $scriptBlock
}

function Remove-HPBloatware-ScriptBlock {
    Write-GuiLog "Removing HP bloatware..." -Color Green
    Write-GuiLog "This can take a while." -Color Yellow
    $scriptBlock = {
        $ErrorActionPreference = 'Continue'
        $AppXPatternsToRemove = @("*HPSupportAssistant*", "*HPJumpStarts*", "*HPPowerManager*", "*HPPrivacySettings*", "*HPSureShield*", "*HPQuickDrop*", "*HPWorkWell*", "*myHP*", "*HPDesktopSupportUtilities*", "*HPQuickTouch*", "*HPEasyClean*", "*HPPCHardwareDiagnosticsWindows*", "*HPSystemInformation*", "AD2F1837.*")
        $ProgramNamePatternsToRemove = @("HP Client Security Manager", "HP Connection Optimizer", "HP Documentation", "HP MAC Address Manager", "HP Notifications", "HP Security Update Service", "HP System Default Settings", "HP Sure Click", "HP Sure Run", "HP Sure Recover", "HP Sure Sense", "HP Wolf Security", "*HP Support Solutions Framework*")

        Microsoft.PowerShell.Host.WriteTranscriptUtil "--- Removing AppX Provisioned Packages (HP) ---"
        Get-AppxProvisionedPackage -Online | Where-Object { $patternMatch = $false; foreach($p in $using:AppXPatternsToRemove){if($_.DisplayName -like $p -or $_.PackageName -like $p){$patternMatch=$true;break}}; $patternMatch } | ForEach-Object { Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing provisioned: $($_.DisplayName)"; Remove-AppxProvisionedPackage -PackageName $_.PackageName -Online -AllUsers -ErrorAction SilentlyContinue }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "--- Removing AppX Packages (HP) ---"
        Get-AppxPackage -AllUsers | Where-Object { $patternMatch = $false; foreach($p in $using:AppXPatternsToRemove){if($_.Name -like $p -or $_.PackageFullName -like $p){$patternMatch=$true;break}}; $patternMatch } | ForEach-Object { Microsoft.PowerShell.Host.WriteTranscriptUtil "Removing AppX: $($_.Name)"; Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction SilentlyContinue }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "--- Uninstalling Programs (HP) ---"
        Get-Package -ErrorAction SilentlyContinue | Where-Object { $patternMatch = $false; foreach($p in $using:ProgramNamePatternsToRemove){if($_.Name -like $p){$patternMatch=$true;break}}; $patternMatch } | ForEach-Object { Microsoft.PowerShell.Host.WriteTranscriptUtil "Uninstalling: $($_.Name)"; try { $_ | Uninstall-Package -Force -ErrorAction Stop } catch { Microsoft.PowerShell.Host.WriteTranscriptUtil "WARN: Failed uninstall: $($_.Name): $($_.Exception.Message)" } }
        
        $MsiExecFallbacks = @('{0E2E04B0-9EDD-11EB-B38C-10604B96B11E}', '{4DA839F0-72CF-11EC-B247-3863BB3CB5A8}')
        foreach ($guid in $MsiExecFallbacks) { Microsoft.PowerShell.Host.WriteTranscriptUtil "MSI uninstall: $guid"; Start-Process -FilePath "msiexec.exe" -ArgumentList "/x `"$guid`" /qn /norestart" -Wait -ErrorAction SilentlyContinue }
        Microsoft.PowerShell.Host.WriteTranscriptUtil "HP bloatware removal done."
    }
    return $scriptBlock
}

# --- GUI Construction ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "RPI Repair & Setup Tool v1.5" # Updated version
$mainForm.Size = New-Object System.Drawing.Size(900, 700); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.MinimumSize = New-Object System.Drawing.Size(750, 550)
$statusStrip = New-Object System.Windows.Forms.StatusStrip; $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel "Ready"; $statusStrip.Items.Add($statusLabel)
$tabControl = New-Object System.Windows.Forms.TabControl; $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$outputPanel = New-Object System.Windows.Forms.Panel; $outputPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom; $outputPanel.Size = New-Object System.Drawing.Size($mainForm.ClientSize.Width, 200); $outputPanel.Controls.Add($script:outputBox)
$mainForm.Controls.Clear(); $mainForm.Controls.AddRange(@($tabControl, $outputPanel, $statusStrip)); $statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom; $outputPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom; $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

function New-ToolButton {
    param([string]$Text, $OnClickActionParam, [System.Windows.Forms.Control]$ParentControl, [string]$JobScriptBlockFunctionNameParam, [string]$OperationNameForJobParam)
    Write-GuiLog "New-ToolButton RECEIVED for '$Text': OnClickSB: $($OnClickActionParam -is [scriptblock]), JobFunc: '$JobScriptBlockFunctionNameParam', OpName: '$OperationNameForJobParam'" -Color Magenta
    $button = New-Object System.Windows.Forms.Button; $button.Text = $Text; $button.AutoSize = $true; $button.MinimumSize = New-Object System.Drawing.Size(250, 0); $button.Padding = New-Object System.Windows.Forms.Padding(10,5,10,5); $button.Margin = New-Object System.Windows.Forms.Padding(5)
    $buttonActionData = @{}
    if ((-not ([string]::IsNullOrWhiteSpace($JobScriptBlockFunctionNameParam))) -and (-not ([string]::IsNullOrWhiteSpace($OperationNameForJobParam)))) {
        $buttonActionData.Type = "Job"; $buttonActionData.JobFuncName = $JobScriptBlockFunctionNameParam; $buttonActionData.OpName = $OperationNameForJobParam
        Write-GuiLog "  CONFIG AS JOB for '$Text'. Func: '$($buttonActionData.JobFuncName)', Op: '$($buttonActionData.OpName)'" -Color DarkCyan
    } elseif ($OnClickActionParam -is [scriptblock]) {
        $buttonActionData.Type = "Direct"; $buttonActionData.Action = $OnClickActionParam
        Write-GuiLog "  CONFIG AS DIRECT for '$Text'. Action type: $($OnClickActionParam.GetType().FullName)" -Color DarkCyan
    } else {
        $button.Text = "$Text (Misconfigured)"; $button.Enabled = $false; Write-GuiLog "Button '$Text' MISCONFIGURED." -Color Red; $ParentControl.Controls.Add($button); return $button
    }
    $button.Tag = $buttonActionData
    $button.Add_Click({
        param($sender, $eventArgs); $clickedButton = $sender -as [System.Windows.Forms.Button]; $retrievedButtonData = $clickedButton.Tag
        Write-GuiLog "CLICKED: '$($clickedButton.Text)'. Tag Data:" -Color Blue; $retrievedButtonData | Format-List | Out-String | ForEach-Object { Write-GuiLog $_ -Color Blue -NoTimestamp }
        if ($null -eq $retrievedButtonData -or $retrievedButtonData.Count -eq 0) { Write-GuiLog "CRITICAL: Tag data NULL/EMPTY for $($clickedButton.Text)!" -Color Red; $statusLabel.Text = "Error: Btn data missing"; return }
        if ($retrievedButtonData.Type -eq "Job") {
            $statusLabel.Text = "Starting job: $($retrievedButtonData.OpName)..."; Write-GuiLog "Preparing job '$($retrievedButtonData.OpName)'. Func from Tag: '$($retrievedButtonData.JobFuncName)'." -Color DarkCyan
            if ([string]::IsNullOrWhiteSpace($retrievedButtonData.JobFuncName)) { Write-GuiLog "CRITICAL: JobFuncName in Tag NULL/EMPTY for $($retrievedButtonData.OpName)." -Color Red; $statusLabel.Text = "Error prep job"; return }
            $scriptBlockFromFunc = $null; try { $scriptBlockFromFunc = Invoke-Expression $retrievedButtonData.JobFuncName -ErrorAction Stop } catch { Write-GuiLog "ERROR invoking '$($retrievedButtonData.JobFuncName)': $($_.Exception.Message)" -Color Red; $statusLabel.Text = "Error invoke func"; return }
            if (-not $scriptBlockFromFunc) { Write-GuiLog "ERROR: Func '$($retrievedButtonData.JobFuncName)' returned NULL." -Color Red; $statusLabel.Text = "Error: Func NULL"; return }
            if ($scriptBlockFromFunc -isnot [scriptblock]) { Write-GuiLog "ERROR: Func '$($retrievedButtonData.JobFuncName)' returned '$($scriptBlockFromFunc.GetType().FullName)' NOT scriptblock." -Color Red; $statusLabel.Text = "Error: Func wrong type"; return }
            Start-LongRunningJob -ScriptBlock $scriptBlockFromFunc -OperationName $retrievedButtonData.OpName -ButtonToDisable $clickedButton
        } elseif ($retrievedButtonData.Type -eq "Direct") {
            $statusLabel.Text = "Executing: $($clickedButton.Text)..."
            try { if ($retrievedButtonData.Action -is [scriptblock]) { & $retrievedButtonData.Action; Write-GuiLog "Direct action '$($clickedButton.Text)' done." -Color Green } else { Write-GuiLog "ERROR: Action in Tag NOT scriptblock. Type: $($retrievedButtonData.Action.GetType().FullName)" -Color Red } }
            catch { Write-GuiLog "Error direct action '$($clickedButton.Text)': $($_.Exception.Message)" -Color Red; if ($retrievedButtonData.Action -is [scriptblock]) { Write-GuiLog "Failing SB: $($retrievedButtonData.Action.ToString())" -Color DarkRed } }
            $statusLabel.Text = "Finished: $($clickedButton.Text)."
        } else { Write-GuiLog "CRITICAL: Unknown Tag type: '$($retrievedButtonData.Type)' for $($clickedButton.Text)" -Color Red; $statusLabel.Text = "Error: Unknown btn type" }
    })
    $ParentControl.Controls.Add($button); return $button
}

function New-ButtonFlowPanel { $panel = New-Object System.Windows.Forms.FlowLayoutPanel; $panel.Dock = [System.Windows.Forms.DockStyle]::Fill; $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown; $panel.AutoScroll = $true; $panel.WrapContents = $false; $panel.Padding = New-Object System.Windows.Forms.Padding(15); return $panel }

# == New PC Setup Tab ==
$tabNewPC = New-Object System.Windows.Forms.TabPage; $tabNewPC.Text = "New PC Setup"; $panelNewPC = New-ButtonFlowPanel; $tabNewPC.Controls.Add($panelNewPC)
New-ToolButton -Text "1: Open New PC Files Folder" -OnClickActionParam {Open-NewPCFiles-Action} -ParentControl $panelNewPC
New-ToolButton -Text "2: Download & Run Ninite" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Download-And-Open-Ninite-ScriptBlock" -OperationNameForJobParam "Ninite Download & Run"
New-ToolButton -Text "3: Download MS Teams (New)" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Download-MS-Teams-ScriptBlock" -OperationNameForJobParam "MS Teams Download"
New-ToolButton -Text "4: Change PC Name (Restarts PC)" -OnClickActionParam {Change-PCName-Action} -ParentControl $panelNewPC
$groupJoinDomain = New-Object System.Windows.Forms.GroupBox; $groupJoinDomain.Text = "5: Join RPI Domain (Restarts PC)"; $groupJoinDomain.AutoSize = $true; $groupJoinDomain.Padding = New-Object System.Windows.Forms.Padding(10); $flowJoinDomain = New-Object System.Windows.Forms.FlowLayoutPanel; $flowJoinDomain.Dock = [System.Windows.Forms.DockStyle]::Fill; $flowJoinDomain.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown; $flowJoinDomain.AutoSize = $true
$lblDomain = New-Object System.Windows.Forms.Label; $lblDomain.Text = "Domain Name (e.g., rootprojects.local):"; $lblDomain.AutoSize=$true; $txtDomainName = New-Object System.Windows.Forms.TextBox; $txtDomainName.Width = 250; $txtDomainName.Text = "rootprojects.local"
$lblOU = New-Object System.Windows.Forms.Label; $lblOU.Text = "OU Path (e.g., OU=Computers,DC=rootprojects,DC=local):"; $lblOU.AutoSize=$true; $txtOUPath = New-Object System.Windows.Forms.TextBox; $txtOUPath.Width = 350; $txtOUPath.Text = "OU=Computers,DC=rootprojects,DC=local"
$btnJoinDomain = New-Object System.Windows.Forms.Button; $btnJoinDomain.Text = "Join Domain"; $btnJoinDomain.Padding = New-Object System.Windows.Forms.Padding(5)
$btnJoinDomain.Add_Click({ $domain = $txtDomainName.Text.Trim(); $ou = $txtOUPath.Text.Trim(); if ([string]::IsNullOrWhiteSpace($domain)) { [System.Windows.Forms.MessageBox]::Show("Domain Name required.", "Error", "OK", "Error"); return }; try { $credential = Get-Credential -Message "Enter credentials for '$domain' domain join"; if ($credential) { Join-Domain-Action -domainName $domain -ouPath $ou -credential $credential } else { Write-GuiLog "Domain join cancelled - no credentials." -Color Yellow } } catch { Write-GuiLog "Domain join failed/cancelled: $($_.Exception.Message)" -Color Yellow } })
$flowJoinDomain.Controls.AddRange(@($lblDomain, $txtDomainName, $lblOU, $txtOUPath, $btnJoinDomain)); $groupJoinDomain.Controls.Add($flowJoinDomain); $panelNewPC.Controls.Add($groupJoinDomain)
New-ToolButton -Text "6: Update Windows (May Restart)" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Update-Windows-ScriptBlock" -OperationNameForJobParam "Windows Update"
New-ToolButton -Text "7: Install Adobe Acrobat Reader" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Install-AdobeReader-ScriptBlock" -OperationNameForJobParam "Adobe Reader Install"
New-ToolButton -Text "8: Remove HP Bloatware" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Remove-HPBloatware-ScriptBlock" -OperationNameForJobParam "HP Bloatware Removal"
$tabControl.Controls.Add($tabNewPC)

# == Windows Repairs Tab ==
$tabWindows = New-Object System.Windows.Forms.TabPage; $tabWindows.Text = "Windows Repairs"; $panelWindows = New-ButtonFlowPanel; $tabWindows.Controls.Add($panelWindows)
New-ToolButton -Text "1: DISM RestoreHealth (May Restart)" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Repair-Windows-ScriptBlock" -OperationNameForJobParam "DISM RestoreHealth"
New-ToolButton -Text "2: System File Checker (SFC)" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Repair-SystemFiles-ScriptBlock" -OperationNameForJobParam "SFC Scan"
New-ToolButton -Text "3: Check Disk (CHKDSK C:) (Restarts)" -OnClickActionParam {Repair-Disk-Action} -ParentControl $panelWindows
New-ToolButton -Text "4: Windows Update Troubleshooter" -OnClickActionParam {Run-WindowsUpdateTroubleshooter-Action} -ParentControl $panelWindows
New-ToolButton -Text "5: DISM Check & Repair (May Restart)" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Check-And-Repair-DISM-ScriptBlock" -OperationNameForJobParam "DISM Check & Repair"
New-ToolButton -Text "6: Network Reset (May Restart)" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Reset-Network-ScriptBlock" -OperationNameForJobParam "Network Reset"
New-ToolButton -Text "7: Windows Memory Diagnostic (Restarts)" -OnClickActionParam {Run-MemoryDiagnostic-Action} -ParentControl $panelWindows
New-ToolButton -Text "8: Windows Startup Repair (Restarts)" -OnClickActionParam {Run-StartupRepair-Action} -ParentControl $panelWindows
New-ToolButton -Text "9: Windows Defender Full Scan" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Run-WindowsDefenderScan-ScriptBlock" -OperationNameForJobParam "Defender Full Scan"
New-ToolButton -Text "10: Reset Windows Update Components" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Reset-WindowsUpdateComponents-ScriptBlock" -OperationNameForJobParam "Reset WU Components"
New-ToolButton -Text "11: List Installed Apps" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "List-InstalledApps-ScriptBlock" -OperationNameForJobParam "List Installed Apps"
New-ToolButton -Text "12: Network Diagnostics" -ParentControl $panelWindows -JobScriptBlockFunctionNameParam "Network-Diagnostics-ScriptBlock" -OperationNameForJobParam "Network Diagnostics"
New-ToolButton -Text "13: Factory Reset (Data Loss & Restarts!)" -OnClickActionParam {Factory-Reset-Action} -ParentControl $panelWindows
$tabControl.Controls.Add($tabWindows)

# == Office Repairs Tab ==
$tabOffice = New-Object System.Windows.Forms.TabPage; $tabOffice.Text = "Office Repairs"; $panelOffice = New-ButtonFlowPanel; $tabOffice.Controls.Add($panelOffice)
New-ToolButton -Text "1: Repair Microsoft Office" -OnClickActionParam {Repair-Office-Action} -ParentControl $panelOffice
New-ToolButton -Text "2: Check for Microsoft Office Updates" -OnClickActionParam {Check-OfficeUpdates-Action} -ParentControl $panelOffice
$tabControl.Controls.Add($tabOffice)

# == User Tasks Tab ==
$tabUserTasks = New-Object System.Windows.Forms.TabPage; $tabUserTasks.Text = "User Tasks"; $panelUserTasks = New-ButtonFlowPanel; $tabUserTasks.Controls.Add($panelUserTasks)
New-ToolButton -Text "1: Clean Temp Files" -ParentControl $panelUserTasks -JobScriptBlockFunctionNameParam "Clean-TempFiles-ScriptBlock" -OperationNameForJobParam "Clean Temp Files"
$groupPrinters = New-Object System.Windows.Forms.GroupBox; $groupPrinters.Text = "2: Printer Mapping"; $groupPrinters.AutoSize = $true; $groupPrinters.Padding = New-Object System.Windows.Forms.Padding(10); $flowPrinters = New-Object System.Windows.Forms.FlowLayoutPanel; $flowPrinters.Dock = [System.Windows.Forms.DockStyle]::Fill; $flowPrinters.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown; $flowPrinters.AutoSize = $true
$printersToMap = @{ "Sydney Printer (192.168.23.10)" = "\\192.168.23.10"; "Melbourne Printer (192.168.33.63)" = "\\192.168.33.63"; "Melb Airport Printer (192.168.43.250)" = "\\192.168.43.250"; "Townsville Printer (192.168.100.240)" = "\\192.168.100.240"; "Brisbane Printer (192.168.20.242)" = "\\192.168.20.242"; "Mackay Printer (192.168.90.240)" = "\\192.168.90.240" }
foreach ($entry in $printersToMap.GetEnumerator()) { $actionScriptBlock = [scriptblock]::Create("Map-Printer-Action -PrinterConnectionName '$($entry.Value)' -PrinterFriendlyName '$($entry.Key)'"); New-ToolButton -Text "Map $($entry.Key)" -OnClickActionParam $actionScriptBlock -ParentControl $flowPrinters }
New-ToolButton -Text "Install All Printers (via VBS)" -OnClickActionParam {Install-AllPrinters-Action} -ParentControl $flowPrinters
$groupPrinters.Controls.Add($flowPrinters); $panelUserTasks.Controls.Add($groupPrinters)
New-ToolButton -Text "3: Clear Teams Cache & Restart Teams" -ParentControl $panelUserTasks -JobScriptBlockFunctionNameParam "Clear-TeamsCache-ScriptBlock" -OperationNameForJobParam "Clear Teams Cache"
New-ToolButton -Text "BONUS: Start Microsoft Teams" -OnClickActionParam {Start-Teams-Action} -ParentControl $panelUserTasks
$tabControl.Controls.Add($tabUserTasks)

# --- Form Load and Show ---
$mainForm.Add_Shown({ Write-GuiLog "RPI Repair & Setup Tool GUI Initialized." -Color Blue; $statusLabel.Text = "Ready." })
$mainForm.Add_FormClosing({ Write-GuiLog "Exiting RPI Repair Tool..." -Color Blue; Get-EventSubscriber -SourceIdentifier JobEvent_* -ErrorAction SilentlyContinue | ForEach-Object { Unregister-Event -SubscriptionId $_.SubscriptionId }; Get-Job | Where-Object {$_.Name -like "*"} | Remove-Job -Force -ErrorAction SilentlyContinue })
[System.Windows.Forms.Application]::EnableVisualStyles(); [void]$mainForm.ShowDialog()
