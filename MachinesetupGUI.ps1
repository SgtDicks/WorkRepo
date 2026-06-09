param(
    [switch]$SmokeTest
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:originalScriptPath = "C:\Users\abycroft\Documents\Codex\2026-05-06\download-the-code-from-https-raw\SetupmachineTest.ps1"
$script:appsPath = "C:\apps"
$script:ariaUrl = "https://github.com/aria2/aria2/releases/download/release-1.37.0/aria2-1.37.0-win-64bit-build1.zip"
$script:ariaZip = "C:\apps\aria2.zip"
$script:ariaFolder = "C:\apps\aria2"
$script:ariaExePath = $null
$script:bluebeamUrl = "https://www.bluebeam.com/MSIdeployx64"
$script:bbOutputFile = "C:\apps\Bluebeam21installer.zip"
$script:bbPath = "C:\apps\Bluebeam21installer"
$script:reviztoUrl = "https://update.revizto.com/v5/msi64"
$script:reviztoMsiPath = "C:\apps\Revizto_x64.msi"
$script:reviztoLogPath = "C:\apps\ReviztoInstall.log"
$script:mainForm = $null
$script:logBox = $null
$script:statusLabel = $null

function Set-Status {
    param(
        [string]$Message
    )

    if ($script:statusLabel -and $script:statusLabel.GetCurrentParent()) {
        if ($script:statusLabel.GetCurrentParent().InvokeRequired) {
            $text = $Message
            $null = $script:statusLabel.GetCurrentParent().BeginInvoke(
                [System.Windows.Forms.MethodInvoker]{
                    Set-Status -Message $text
                }
            )
            return
        }

        $script:statusLabel.Text = $Message
        $script:statusLabel.GetCurrentParent().Refresh()
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Level = "Info"
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    if (-not $script:logBox) {
        Write-Output "[$Level] $Message"
        return
    }

    if ($script:logBox.InvokeRequired) {
        $text = $Message
        $severity = $Level
        $null = $script:logBox.BeginInvoke(
            [System.Windows.Forms.MethodInvoker]{
                Write-Log -Message $text -Level $severity
            }
        )
        return
    }

    $timestamp = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) {
        "Success" { [System.Drawing.Color]::LightGreen; break }
        "Warning" { [System.Drawing.Color]::Khaki; break }
        "Error" { [System.Drawing.Color]::LightCoral; break }
        default { [System.Drawing.Color]::Gainsboro }
    }

    $script:logBox.SelectionStart = $script:logBox.TextLength
    $script:logBox.SelectionLength = 0
    $script:logBox.SelectionColor = [System.Drawing.Color]::SkyBlue
    $script:logBox.AppendText("[$timestamp] ")
    $script:logBox.SelectionColor = $color
    $script:logBox.AppendText("$Message`r`n")
    $script:logBox.SelectionColor = $script:logBox.ForeColor
    $script:logBox.ScrollToCaret()
    $script:logBox.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Write-TextBlock {
    param(
        [string]$Text,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Level = "Info"
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    foreach ($line in ($Text -split "(`r`n|`n|`r)")) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            Write-Log -Message $line.TrimEnd() -Level $Level
        }
    }
}

function Show-ErrorDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    [System.Windows.Forms.MessageBox]::Show(
        $script:mainForm,
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Confirm-Action {
    param(
        [string]$Message
    )

    $result = [System.Windows.Forms.MessageBox]::Show(
        $script:mainForm,
        "$Message`r`n`r`nDo you want to proceed?",
        "Confirm Action",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        return $true
    }

    Write-Log -Message "Operation canceled by user." -Level "Warning"
    return $false
}

function Select-FolderPath {
    param(
        [string]$Description
    )

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.UseDescriptionForTitle = $true
    $dialog.ShowNewFolderButton = $true

    try {
        if ($dialog.ShowDialog($script:mainForm) -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
    } finally {
        $dialog.Dispose()
    }

    return $null
}

function Invoke-UiAction {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    try {
        Set-Status -Message "Running: $Name"
        $script:mainForm.UseWaitCursor = $true
        $script:mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Write-Log -Message "Starting $Name." -Level "Info"
        [System.Windows.Forms.Application]::DoEvents()
        & $Action
        Write-Log -Message "$Name finished." -Level "Success"
        Set-Status -Message "Ready"
    } catch {
        $message = $_.Exception.Message
        Write-Log -Message "$Name failed. $message" -Level "Error"
        Set-Status -Message "Failed: $Name"
        Show-ErrorDialog -Title $Name -Message $message
    } finally {
        $script:mainForm.UseWaitCursor = $false
        $script:mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Ensure-Directory {
    param(
        [string]$Path
    )

    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Log -Message "Created directory: $Path" -Level "Success"
    }
}

function Wait-ForProcess {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$Description
    )

    while (-not $Process.HasExited) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
    }

    $Process.WaitForExit()

    if ($Process.ExitCode -ne 0) {
        throw "$Description exited with code $($Process.ExitCode)."
    }
}

function Start-LoggedProcess {
    param(
        [string]$FilePath,
        [string]$Arguments,
        [string]$Description,
        [switch]$Wait,
        [switch]$RunAs
    )

    $label = if ($Description) { $Description } else { Split-Path -Path $FilePath -Leaf }
    Write-Log -Message "Launching $label." -Level "Info"

    $processArgs = @{
        FilePath = $FilePath
    }

    if ($Arguments) {
        $processArgs.ArgumentList = $Arguments
    }

    if ($RunAs) {
        $processArgs.Verb = "RunAs"
    }

    if ($Wait) {
        $processArgs.PassThru = $true
        $processArgs.Wait = $true
    }

    $process = Start-Process @processArgs

    if ($Wait -and $process.ExitCode -ne 0) {
        throw "$label exited with code $($process.ExitCode)."
    }
}

function Invoke-ConsoleCommand {
    param(
        [string]$FilePath,
        [string]$Arguments,
        [string]$Description
    )

    Write-Log -Message "Running $Description." -Level "Info"

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.Arguments = $Arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $process.EnableRaisingEvents = $true

    $process.add_OutputDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data) {
            Write-Log -Message $eventArgs.Data -Level "Info"
        }
    })

    $process.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($eventArgs.Data) {
            Write-Log -Message $eventArgs.Data -Level "Error"
        }
    })

    [void]$process.Start()
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()

    Wait-ForProcess -Process $process -Description $Description
}

function Ensure-Aria2Ready {
    Ensure-Directory -Path $script:appsPath
    Ensure-Directory -Path $script:ariaFolder

    $script:ariaExePath = Get-ChildItem -Path $script:ariaFolder -Recurse -Filter "aria2c.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($script:ariaExePath) {
        Write-Log -Message "aria2 is ready: $($script:ariaExePath.FullName)" -Level "Success"
        return
    }

    Write-Log -Message "Downloading aria2 support files." -Level "Info"
    Invoke-WebRequest -Uri $script:ariaUrl -OutFile $script:ariaZip
    Write-Log -Message "Extracting aria2 support files." -Level "Info"
    Expand-Archive -Path $script:ariaZip -DestinationPath $script:ariaFolder -Force

    $script:ariaExePath = Get-ChildItem -Path $script:ariaFolder -Recurse -Filter "aria2c.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $script:ariaExePath) {
        throw "aria2c.exe not found after extraction."
    }

    Write-Log -Message "aria2 download helper is ready." -Level "Success"
}

function Initialize-SupportFiles {
    Ensure-Directory -Path $script:appsPath
    Ensure-Directory -Path $script:bbPath
    Write-Log -Message "GUI is ready. Opened from $script:originalScriptPath" -Level "Success"
}

function Change-PCName {
    Write-Log -Message "Changing the PC name based on the serial number." -Level "Info"
    $serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber
    $newPCName = "RPI-$serialNumber"
    Write-Log -Message "The new PC name will be: $newPCName" -Level "Warning"

    if (Confirm-Action -Message "This will rename the computer to $newPCName and require a restart.") {
        Rename-Computer -NewName $newPCName -Force -Restart
    }
}

function Join-Domain {
    param(
        [string]$OU,
        [switch]$Restart,
        [switch]$Force
    )

    $domainName = "rootprojects.local"
    if ((Get-CimInstance Win32_ComputerSystem).PartOfDomain) {
        Write-Log -Message "This machine is already joined to a domain." -Level "Warning"
        return
    }

    $credential = Get-Credential -Message "Enter credentials to join $domainName"
    $parameters = @{
        DomainName = $domainName
        Credential = $credential
    }

    if ($OU) {
        $parameters.OUPath = $OU
    }

    if ($Restart) {
        $parameters.Restart = $true
    }

    if ($Force) {
        $parameters.Force = $true
    }

    Add-Computer @parameters
    Write-Log -Message "Domain join command completed." -Level "Success"
}

function Repair-Windows {
    Write-Log -Message "This process can take 15-30 minutes and may require a restart." -Level "Warning"
    if (Confirm-Action -Message "Run DISM RestoreHealth?") {
        Invoke-ConsoleCommand -FilePath "dism.exe" -Arguments "/online /cleanup-image /restorehealth" -Description "DISM RestoreHealth"
    }
}

function Repair-SystemFiles {
    Write-Log -Message "This scan can take up to 20 minutes." -Level "Warning"
    if (Confirm-Action -Message "Run System File Checker?") {
        Invoke-ConsoleCommand -FilePath "sfc.exe" -Arguments "/scannow" -Description "System File Checker"
    }
}

function Repair-Disk {
    Write-Log -Message "This operation may take several hours and will restart the computer." -Level "Warning"
    if (Confirm-Action -Message "Run CHKDSK on C:?") {
        Invoke-ConsoleCommand -FilePath "chkdsk.exe" -Arguments "C: /F /R /X" -Description "CHKDSK"
    }
}

function Run-WindowsUpdateTroubleshooter {
    Write-Log -Message "This operation may take 5-10 minutes." -Level "Info"
    if (Confirm-Action -Message "Run Windows Update Troubleshooter?") {
        Start-LoggedProcess -FilePath "msdt.exe" -Arguments "/id WindowsUpdateDiagnostic" -Description "Windows Update Troubleshooter" -Wait
    }
}

function Check-And-Repair-DISM {
    Write-Log -Message "This process might take 30-60 minutes." -Level "Warning"
    if (Confirm-Action -Message "Run DISM CheckHealth, ScanHealth, and RestoreHealth?") {
        Invoke-ConsoleCommand -FilePath "dism.exe" -Arguments "/Online /Cleanup-Image /CheckHealth" -Description "DISM CheckHealth"
        Invoke-ConsoleCommand -FilePath "dism.exe" -Arguments "/Online /Cleanup-Image /ScanHealth" -Description "DISM ScanHealth"
        Invoke-ConsoleCommand -FilePath "dism.exe" -Arguments "/Online /Cleanup-Image /RestoreHealth" -Description "DISM RestoreHealth"
    }
}

function Reset-Network {
    Write-Log -Message "This will temporarily disrupt network connectivity." -Level "Warning"
    if (Confirm-Action -Message "Reset the network adapters and IP stack?") {
        Invoke-ConsoleCommand -FilePath "netsh.exe" -Arguments "winsock reset" -Description "Winsock reset"
        Invoke-ConsoleCommand -FilePath "netsh.exe" -Arguments "int ip reset" -Description "TCP/IP reset"
        Invoke-ConsoleCommand -FilePath "ipconfig.exe" -Arguments "/release" -Description "IP release"
        Invoke-ConsoleCommand -FilePath "ipconfig.exe" -Arguments "/renew" -Description "IP renew"
        Invoke-ConsoleCommand -FilePath "ipconfig.exe" -Arguments "/flushdns" -Description "DNS flush"
        Write-Log -Message "Network reset completed. A restart may be required." -Level "Success"
    }
}

function Run-MemoryDiagnostic {
    Write-Log -Message "This test will restart the computer." -Level "Warning"
    if (Confirm-Action -Message "Schedule Windows Memory Diagnostic?") {
        Start-LoggedProcess -FilePath "mdsched.exe" -Arguments "/f" -Description "Windows Memory Diagnostic" -RunAs
    }
}

function Run-StartupRepair {
    Write-Log -Message "This process will restart the computer and attempt to fix startup issues." -Level "Warning"
    if (Confirm-Action -Message "Boot into Windows Startup Repair?") {
        Start-LoggedProcess -FilePath "reagentc.exe" -Arguments "/boottore" -Description "Startup Repair preparation" -RunAs -Wait
        Invoke-ConsoleCommand -FilePath "shutdown.exe" -Arguments "/r /t 0" -Description "Immediate restart"
    }
}

function Run-WindowsDefenderScan {
    Write-Log -Message "This scan can take several hours." -Level "Warning"
    if (Confirm-Action -Message "Run a Windows Defender Full Scan?") {
        Start-MpScan -ScanType FullScan
        Write-Log -Message "Windows Defender Full Scan completed." -Level "Success"
    }
}

function Reset-WindowsUpdateComponents {
    Write-Log -Message "Windows Update services will be temporarily unavailable." -Level "Warning"
    if (Confirm-Action -Message "Reset Windows Update components?") {
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "stop wuauserv" -Description "Stop Windows Update service"
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "stop cryptsvc" -Description "Stop Cryptographic Services"
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "stop bits" -Description "Stop BITS"
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "stop msiserver" -Description "Stop Windows Installer service"
        Rename-Item -Path "C:\Windows\SoftwareDistribution" -NewName "SoftwareDistribution.old" -Force -ErrorAction SilentlyContinue
        Rename-Item -Path "C:\Windows\System32\catroot2" -NewName "Catroot2.old" -Force -ErrorAction SilentlyContinue
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "start wuauserv" -Description "Start Windows Update service"
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "start cryptsvc" -Description "Start Cryptographic Services"
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "start bits" -Description "Start BITS"
        Invoke-ConsoleCommand -FilePath "net.exe" -Arguments "start msiserver" -Description "Start Windows Installer service"
        Write-Log -Message "Windows Update components reset completed." -Level "Success"
    }
}

function Start-Teams {
    Write-Log -Message "Trying to find and launch Microsoft Teams." -Level "Info"

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
        $match = Get-Item -Path $path -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($match) {
            $teamsPath = $match.FullName
            Write-Log -Message "Teams executable found at: $teamsPath" -Level "Success"
            break
        }

        Write-Log -Message "Teams executable not found at: $path" -Level "Warning"
    }

    if ($teamsPath) {
        Start-LoggedProcess -FilePath $teamsPath -Description "Microsoft Teams"
    } else {
        throw "Microsoft Teams executable not found."
    }
}

function Clear-TeamsCache {
    if (-not (Confirm-Action -Message "Delete the Teams cache?")) {
        return
    }

    Write-Log -Message "Closing Teams." -Level "Info"

    foreach ($processName in @("ms-teams", "Teams")) {
        try {
            $teamProcesses = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($teamProcesses) {
                $teamProcesses | Stop-Process -Force
                Start-Sleep -Seconds 3
                Write-Log -Message "Closed process: $processName" -Level "Success"
            }
        } catch {
            Write-Log -Message "Failed to close process $processName. $($_.Exception.Message)" -Level "Warning"
        }
    }

    $cachePath = "$env:LOCALAPPDATA\Packages\MSTeams_*\LocalCache\Local\Microsoft\Teams"
    $cacheItems = Get-Item -Path $cachePath -ErrorAction SilentlyContinue

    if (-not $cacheItems) {
        Write-Log -Message "Teams cache path not found: $cachePath" -Level "Warning"
        return
    }

    foreach ($cacheItem in $cacheItems) {
        Remove-Item -Path $cacheItem.FullName -Recurse -Force -Confirm:$false
        Write-Log -Message "Removed Teams cache: $($cacheItem.FullName)" -Level "Success"
    }

    Write-Log -Message "Cleanup complete. Trying to launch Teams." -Level "Success"
    Start-Teams
}

function Repair-Office {
    $officeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (-not (Test-Path $officeClickToRunPath)) {
        throw "Microsoft Office Click-to-Run client not found. Please ensure Office is installed."
    }

    Start-LoggedProcess -FilePath $officeClickToRunPath -Arguments "scenario=Repair" -Description "Microsoft Office repair" -Wait
}

function Check-OfficeUpdates {
    $officeClickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (-not (Test-Path $officeClickToRunPath)) {
        throw "Microsoft Office Click-to-Run client not found. Please ensure Office is installed."
    }

    Start-LoggedProcess -FilePath $officeClickToRunPath -Arguments "scenario=ApplyUpdates" -Description "Office updates" -Wait
}

function Factory-Reset {
    Write-Log -Message "WARNING: All personal files, apps, and settings will be removed." -Level "Error"
    if (Confirm-Action -Message "Initiate a factory reset?") {
        Start-LoggedProcess -FilePath "systemreset.exe" -Arguments "-factoryreset" -Description "Factory Reset" -RunAs
    }
}

function Update-Windows {
    Write-Log -Message "Checking for Windows updates." -Level "Info"

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log -Message "PSWindowsUpdate module not found. Installing." -Level "Warning"
        Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force
        Write-Log -Message "PSWindowsUpdate module installed successfully." -Level "Success"
    }

    Import-Module PSWindowsUpdate
    $updateOutput = Install-WindowsUpdate -AcceptAll -AutoReboot | Out-String
    Write-TextBlock -Text $updateOutput -Level "Info"
    Write-Log -Message "Windows update command completed." -Level "Success"
}

function List-InstalledApps {
    Write-Log -Message "Listing installed applications. This may take some time." -Level "Info"
    $output = Get-WmiObject -Class Win32_Product | Select-Object Name, Version | Format-Table -AutoSize | Out-String
    Write-TextBlock -Text $output -Level "Info"
}

function Network-Diagnostics {
    Write-Log -Message "Performing network diagnostics." -Level "Info"

    $googleOutput = Test-Connection -ComputerName "google.com" -Count 4 | Format-Table -AutoSize | Out-String
    Write-TextBlock -Text $googleOutput -Level "Info"

    foreach ($server in @("10.60.70.11", "192.168.20.186")) {
        Write-Log -Message "Pinging local server $server." -Level "Info"
        $serverOutput = Test-Connection -ComputerName $server -Count 4 | Format-Table -AutoSize | Out-String
        Write-TextBlock -Text $serverOutput -Level "Info"
    }

    Write-Log -Message "Network diagnostics completed." -Level "Success"
}

function Map-Printer {
    param(
        [string]$PrinterIP
    )

    $printerName = "\\$PrinterIP"
    Write-Log -Message "Mapping printer at $PrinterIP." -Level "Info"
    Add-Printer -ConnectionName $printerName
    Write-Log -Message "Printer mapped successfully." -Level "Success"
}

function Open-NewPCFiles {
    Write-Log -Message "Opening New PC Files folder." -Level "Info"
    Start-LoggedProcess -FilePath "explorer.exe" -Arguments "\\RPI-AUS-FS01.rootprojects.local\RPI Admin Archive\Software\RP Files" -Description "New PC Files folder"
}

function Download-And-Open-Ninite {
    Ensure-Directory -Path $script:appsPath
    $niniteUrl = "https://ninite.com/.net4.8-.net4.8.1-7zip-chrome-vlc-zoom/ninite.exe"
    $outputPath = Join-Path -Path $script:appsPath -ChildPath "ninite.exe"

    Write-Log -Message "Downloading Ninite installer." -Level "Info"
    Invoke-WebRequest -Uri $niniteUrl -OutFile $outputPath
    Write-Log -Message "Downloaded Ninite to $outputPath" -Level "Success"
    Start-LoggedProcess -FilePath $outputPath -Description "Ninite installer"
}

function Download-MS-Teams {
    Ensure-Aria2Ready

    $teamsBootstrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $teamsMsixUrl = "https://go.microsoft.com/fwlink/?linkid=2196106"
    $bootstrapperPath = Join-Path -Path $script:appsPath -ChildPath "Teams_bootstrapper.exe"
    $msixPath = Join-Path -Path $script:appsPath -ChildPath "teams.msix"

    Write-Log -Message "Downloading Teams Bootstrapper." -Level "Info"
    Invoke-WebRequest -Uri $teamsBootstrapperUrl -OutFile $bootstrapperPath
    Write-Log -Message "Downloaded Teams Bootstrapper to $bootstrapperPath" -Level "Success"

    Write-Log -Message "Downloading Teams MSIX package." -Level "Info"
    $teamsArgs = "-x 1 -s 1 -o `"$([System.IO.Path]::GetFileName($msixPath))`" -d `"$([System.IO.Path]::GetDirectoryName($msixPath))`" $teamsMsixUrl"
    Start-LoggedProcess -FilePath $script:ariaExePath.FullName -Arguments $teamsArgs -Description "Teams MSIX download" -Wait
    Write-Log -Message "Downloaded Teams MSIX package to $msixPath" -Level "Success"

    Write-Log -Message "Waiting for 5 seconds before installation." -Level "Info"
    Start-Sleep -Seconds 5

    Start-LoggedProcess -FilePath $bootstrapperPath -Arguments "-p -o `"$msixPath`"" -Description "Teams Bootstrapper" -Wait
}

function Download-Agent {
    Ensure-Directory -Path $script:appsPath
    $agentUrl = "https://setup.auplatform.connectwise.com/windows/BareboneAgent/32/Main-RP_Infrastructure_Pty_Ltd_Windows_OS_ITSPlatform_TKNe0edb98f-608d-481f-99a3-8bb6465a4f61/MSI/setup"
    $agentPath = Join-Path -Path $script:appsPath -ChildPath "Main-RP_Infrastructure_Pty_Ltd_Windows_OS_ITSPlatform_TKNe0edb98f-608d-481f-99a3-8bb6465a4f61.msi"

    Write-Log -Message "Downloading First Focus Agent." -Level "Info"
    Invoke-WebRequest -Uri $agentUrl -OutFile $agentPath
    Write-Log -Message "Downloaded Agent MSI package to $agentPath" -Level "Success"
    Start-LoggedProcess -FilePath $agentPath -Description "First Focus Agent installer" -Wait
}

function Download-Bluebeam21 {
    Ensure-Aria2Ready
    Ensure-Directory -Path $script:bbPath

    Write-Log -Message "Downloading Bluebeam 21 installer package." -Level "Info"
    $bluebeamArgs = "-x 16 -s 16 -o `"$([System.IO.Path]::GetFileName($script:bbOutputFile))`" -d `"$([System.IO.Path]::GetDirectoryName($script:bbOutputFile))`" $script:bluebeamUrl"
    Start-LoggedProcess -FilePath $script:ariaExePath.FullName -Arguments $bluebeamArgs -Description "Bluebeam 21 download" -Wait

    Write-Log -Message "Extracting Bluebeam installer." -Level "Info"
    Expand-Archive -Path $script:bbOutputFile -DestinationPath $script:bbPath -Force

    $bbMsi = Get-ChildItem -Path $script:bbPath -Filter "*.msi" -Recurse | Select-Object -First 1
    if (-not $bbMsi) {
        throw "MSI not found inside $($script:bbPath)."
    }

    Write-Log -Message "Running $($bbMsi.Name)." -Level "Info"
    Start-LoggedProcess -FilePath $bbMsi.FullName -Description "Bluebeam 21 installer" -Wait
}


function Download-Revizto {
    Ensure-Aria2Ready

    Write-Log -Message "This will download and silently install Revizto." -Level "Warning"
    if (-not (Confirm-Action -Message "Install Revizto?")) {
        return
    }

    if (Test-Path -Path $script:reviztoMsiPath) {
        Write-Log -Message "Removing existing Revizto installer: $($script:reviztoMsiPath)" -Level "Info"
        Remove-Item -Path $script:reviztoMsiPath -Force
    }

    Write-Log -Message "Downloading Revizto installer with aria2." -Level "Info"
    $reviztoArgs = "-x 16 -s 16 -o `"$([System.IO.Path]::GetFileName($script:reviztoMsiPath))`" -d `"$([System.IO.Path]::GetDirectoryName($script:reviztoMsiPath))`" $script:reviztoUrl"
    Start-LoggedProcess -FilePath $script:ariaExePath.FullName -Arguments $reviztoArgs -Description "Revizto download" -Wait

    if (-not (Test-Path -Path $script:reviztoMsiPath)) {
        throw "Revizto MSI was not downloaded to $($script:reviztoMsiPath)."
    }

    Write-Log -Message "Downloaded Revizto MSI package to $($script:reviztoMsiPath)" -Level "Success"
    Write-Log -Message "Installing Revizto silently. Log file: $($script:reviztoLogPath)" -Level "Info"

    $arguments = "/i `"$($script:reviztoMsiPath)`" /qn /norestart /l*v `"$($script:reviztoLogPath)`""
    Start-LoggedProcess -FilePath "msiexec.exe" -Arguments $arguments -Description "Revizto installer" -Wait

    Write-Log -Message "Revizto installation completed." -Level "Success"
}

function Install-AdobeReader {
    Write-Log -Message "This will download and install Adobe Acrobat Reader 32-bit." -Level "Warning"
    if (-not (Confirm-Action -Message "Install Adobe Acrobat Reader 32-bit?")) {
        return
    }

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "winget is not installed or not found in PATH."
    }

    Invoke-ConsoleCommand -FilePath "winget.exe" -Arguments "install -e --id Adobe.Acrobat.Reader.32-bit -h" -Description "Adobe Acrobat Reader installation"
}

function Remove-HPBloatware {
    if (-not (Confirm-Action -Message "Proceed with removing HP bloatware?")) {
        return
    }

    $uninstallPackages = @(
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

    $uninstallPrograms = @(
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

    $hpIdentifier = "AD2F1837"
    $installedPackages = Get-AppxPackage -AllUsers | Where-Object { ($uninstallPackages -contains $_.Name) -or ($_.Name -match "^$hpIdentifier") }
    $provisionedPackages = Get-AppxProvisionedPackage -Online | Where-Object { ($uninstallPackages -contains $_.DisplayName) -or ($_.DisplayName -match "^$hpIdentifier") }
    $installedPrograms = Get-Package | Where-Object { $uninstallPrograms -contains $_.Name }

    foreach ($provPackage in $provisionedPackages) {
        Write-Log -Message "Attempting to remove provisioned package: [$($provPackage.DisplayName)]" -Level "Info"
        try {
            Remove-AppxProvisionedPackage -PackageName $provPackage.PackageName -Online -ErrorAction Stop | Out-Null
            Write-Log -Message "Successfully removed provisioned package: [$($provPackage.DisplayName)]" -Level "Success"
        } catch {
            Write-Log -Message "Failed to remove provisioned package: [$($provPackage.DisplayName)]" -Level "Warning"
        }
    }

    foreach ($appxPackage in $installedPackages) {
        Write-Log -Message "Attempting to remove Appx package: [$($appxPackage.Name)]" -Level "Info"
        try {
            Remove-AppxPackage -Package $appxPackage.PackageFullName -AllUsers -ErrorAction Stop | Out-Null
            Write-Log -Message "Successfully removed Appx package: [$($appxPackage.Name)]" -Level "Success"
        } catch {
            Write-Log -Message "Failed to remove Appx package: [$($appxPackage.Name)]" -Level "Warning"
        }
    }

    foreach ($program in $installedPrograms) {
        Write-Log -Message "Attempting to uninstall: [$($program.Name)]" -Level "Info"
        try {
            $program | Uninstall-Package -AllVersions -Force -ErrorAction Stop | Out-Null
            Write-Log -Message "Successfully uninstalled: [$($program.Name)]" -Level "Success"
        } catch {
            Write-Log -Message "Failed to uninstall: [$($program.Name)]" -Level "Warning"
        }
    }

    foreach ($productCode in @(
        "{0E2E04B0-9EDD-11EB-B38C-10604B96B11E}",
        "{4DA839F0-72CF-11EC-B247-3863BB3CB5A8}"
    )) {
        try {
            Start-LoggedProcess -FilePath "msiexec.exe" -Arguments "/x `"$productCode`" /qn /norestart" -Description "HP Wolf Security uninstall" -Wait
        } catch {
            Write-Log -Message "Failed MSI uninstall for product $productCode. $($_.Exception.Message)" -Level "Warning"
        }
    }

    Write-Log -Message "HP bloatware removal process completed." -Level "Success"
}

function Install-AllPrinters {
    Write-Log -Message "Installing all printers via VBS script." -Level "Info"

    $vbsPaths = @(
        "\\RPI-AUS-FS01.rootprojects.local\RPI Admin Archive\Software\RP Files\Printers.vbs"
    )

    foreach ($path in $vbsPaths) {
        if (Test-Path $path) {
            Write-Log -Message "Found Printers.vbs at $path" -Level "Success"
            Start-LoggedProcess -FilePath "wscript.exe" -Arguments "`"$path`"" -Description "Printers.vbs" -Wait
            return
        }

        Write-Log -Message "Printers.vbs not found at $path" -Level "Warning"
    }

    throw "Printers.vbs script not found in any of the specified locations."
}

function Clean-TempFiles {
    Write-Log -Message "Cleaning temporary files." -Level "Info"
    foreach ($path in @("$env:TEMP\*", "C:\Windows\Temp\*")) {
        try {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log -Message "Cleared: $path" -Level "Success"
        } catch {
            Write-Log -Message "Failed to clear: $path. $($_.Exception.Message)" -Level "Warning"
        }
    }
}

function New-ActionButton {
    param(
        [string]$Text,
        [scriptblock]$OnClick
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Width = 330
    $button.Height = 38
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.BackColor = [System.Drawing.Color]::FromArgb(34, 93, 138)
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.Add_Click($OnClick)
    return $button
}

function New-SectionLabel {
    param(
        [string]$Text
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 8)
    $label.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
    return $label
}

function New-TabFlowPanel {
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $panel.WrapContents = $false
    $panel.AutoScroll = $true
    $panel.Padding = New-Object System.Windows.Forms.Padding(18, 12, 18, 12)
    $panel.BackColor = [System.Drawing.Color]::WhiteSmoke
    return $panel
}

$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Text = "RPI Repair Menu - GUI"
$script:mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$script:mainForm.Size = New-Object System.Drawing.Size(1220, 760)
$script:mainForm.MinimumSize = New-Object System.Drawing.Size(1040, 640)
$script:mainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height = 78
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(23, 41, 64)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "RPI Repair Menu"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 18)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(18, 14)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "GUI wrapper for the existing PowerShell repair and setup actions."
$subtitleLabel.ForeColor = [System.Drawing.Color]::LightSteelBlue
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$subtitleLabel.AutoSize = $true
$subtitleLabel.Location = New-Object System.Drawing.Point(20, 46)

$headerPanel.Controls.Add($titleLabel)
$headerPanel.Controls.Add($subtitleLabel)

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContainer.SplitterDistance = 690
$splitContainer.BackColor = [System.Drawing.Color]::FromArgb(220, 224, 230)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabs.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$windowsTab = New-Object System.Windows.Forms.TabPage
$windowsTab.Text = "Windows Repairs"
$windowsFlow = New-TabFlowPanel
$windowsFlow.Controls.Add((New-SectionLabel -Text "Repair and recovery"))
$windowsFlow.Controls.Add((New-ActionButton -Text "DISM RestoreHealth" -OnClick { Invoke-UiAction -Name "DISM RestoreHealth" -Action { Repair-Windows } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "System File Checker (SFC)" -OnClick { Invoke-UiAction -Name "System File Checker" -Action { Repair-SystemFiles } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Check Disk (CHKDSK)" -OnClick { Invoke-UiAction -Name "Check Disk" -Action { Repair-Disk } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Windows Update Troubleshooter" -OnClick { Invoke-UiAction -Name "Windows Update Troubleshooter" -Action { Run-WindowsUpdateTroubleshooter } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "DISM Check and Repair" -OnClick { Invoke-UiAction -Name "DISM Check and Repair" -Action { Check-And-Repair-DISM } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Network Reset" -OnClick { Invoke-UiAction -Name "Network Reset" -Action { Reset-Network } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Windows Memory Diagnostic" -OnClick { Invoke-UiAction -Name "Windows Memory Diagnostic" -Action { Run-MemoryDiagnostic } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Windows Startup Repair" -OnClick { Invoke-UiAction -Name "Windows Startup Repair" -Action { Run-StartupRepair } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Windows Defender Full Scan" -OnClick { Invoke-UiAction -Name "Windows Defender Full Scan" -Action { Run-WindowsDefenderScan } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Reset Windows Update Components" -OnClick { Invoke-UiAction -Name "Reset Windows Update Components" -Action { Reset-WindowsUpdateComponents } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "List Installed Apps" -OnClick { Invoke-UiAction -Name "List Installed Apps" -Action { List-InstalledApps } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Network Diagnostics" -OnClick { Invoke-UiAction -Name "Network Diagnostics" -Action { Network-Diagnostics } }))
$windowsFlow.Controls.Add((New-ActionButton -Text "Factory Reset Device / Reinstall Windows" -OnClick { Invoke-UiAction -Name "Factory Reset" -Action { Factory-Reset } }))
$windowsTab.Controls.Add($windowsFlow)

$officeTab = New-Object System.Windows.Forms.TabPage
$officeTab.Text = "Office Repairs"
$officeFlow = New-TabFlowPanel
$officeFlow.Controls.Add((New-SectionLabel -Text "Microsoft Office"))
$officeFlow.Controls.Add((New-ActionButton -Text "Repair Microsoft Office" -OnClick { Invoke-UiAction -Name "Repair Microsoft Office" -Action { Repair-Office } }))
$officeFlow.Controls.Add((New-ActionButton -Text "Check for Microsoft Office Updates" -OnClick { Invoke-UiAction -Name "Check Office Updates" -Action { Check-OfficeUpdates } }))
$officeTab.Controls.Add($officeFlow)

$userTasksTab = New-Object System.Windows.Forms.TabPage
$userTasksTab.Text = "User Tasks"
$userFlow = New-TabFlowPanel
$userFlow.Controls.Add((New-SectionLabel -Text "Common tasks"))
$userFlow.Controls.Add((New-ActionButton -Text "Clean Temp Files" -OnClick { Invoke-UiAction -Name "Clean Temp Files" -Action { Clean-TempFiles } }))
$userFlow.Controls.Add((New-ActionButton -Text "Clear Teams Cache" -OnClick { Invoke-UiAction -Name "Clear Teams Cache" -Action { Clear-TeamsCache } }))
$userFlow.Controls.Add((New-SectionLabel -Text "Printer mapping"))
$userFlow.Controls.Add((New-ActionButton -Text "Map Sydney Printer" -OnClick { Invoke-UiAction -Name "Map Sydney Printer" -Action { Map-Printer -PrinterIP "192.168.23.10" } }))
$userFlow.Controls.Add((New-ActionButton -Text "Map Melbourne Printer" -OnClick { Invoke-UiAction -Name "Map Melbourne Printer" -Action { Map-Printer -PrinterIP "192.168.33.63" } }))
$userFlow.Controls.Add((New-ActionButton -Text "Map Melbourne Airport Printer" -OnClick { Invoke-UiAction -Name "Map Melbourne Airport Printer" -Action { Map-Printer -PrinterIP "192.168.43.250" } }))
$userFlow.Controls.Add((New-ActionButton -Text "Map Townsville Printer" -OnClick { Invoke-UiAction -Name "Map Townsville Printer" -Action { Map-Printer -PrinterIP "192.168.100.240" } }))
$userFlow.Controls.Add((New-ActionButton -Text "Map Brisbane Printer" -OnClick { Invoke-UiAction -Name "Map Brisbane Printer" -Action { Map-Printer -PrinterIP "192.168.20.242" } }))
$userFlow.Controls.Add((New-ActionButton -Text "Map Mackay Printer" -OnClick { Invoke-UiAction -Name "Map Mackay Printer" -Action { Map-Printer -PrinterIP "192.168.90.240" } }))
$userFlow.Controls.Add((New-ActionButton -Text "Install All Printers" -OnClick { Invoke-UiAction -Name "Install All Printers" -Action { Install-AllPrinters } }))
$userTasksTab.Controls.Add($userFlow)

$newPcTab = New-Object System.Windows.Forms.TabPage
$newPcTab.Text = "New PC Setup"
$newPcFlow = New-TabFlowPanel
$newPcFlow.Controls.Add((New-SectionLabel -Text "Provisioning and setup"))
$newPcFlow.Controls.Add((New-ActionButton -Text "Open New PC Files Folder" -OnClick { Invoke-UiAction -Name "Open New PC Files Folder" -Action { Open-NewPCFiles } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Download and Open Ninite" -OnClick { Invoke-UiAction -Name "Download and Open Ninite" -Action { Download-And-Open-Ninite } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Download MS Teams" -OnClick { Invoke-UiAction -Name "Download MS Teams" -Action { Download-MS-Teams } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Change PC Name" -OnClick { Invoke-UiAction -Name "Change PC Name" -Action { Change-PCName } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Join RPI Domain" -OnClick { Invoke-UiAction -Name "Join RPI Domain" -Action { Join-Domain } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Update Windows" -OnClick { Invoke-UiAction -Name "Update Windows" -Action { Update-Windows } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Install Adobe Acrobat Reader 32-bit" -OnClick { Invoke-UiAction -Name "Install Adobe Reader" -Action { Install-AdobeReader } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Remove HP Bloatware" -OnClick { Invoke-UiAction -Name "Remove HP Bloatware" -Action { Remove-HPBloatware } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Install First Focus Agent" -OnClick { Invoke-UiAction -Name "Install First Focus Agent" -Action { Download-Agent } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Install Bluebeam 21" -OnClick { Invoke-UiAction -Name "Install Bluebeam 21" -Action { Download-Bluebeam21 } }))
$newPcFlow.Controls.Add((New-ActionButton -Text "Install Revizto" -OnClick { Invoke-UiAction -Name "Install Revizto" -Action { Download-Revizto } }))
$newPcTab.Controls.Add($newPcFlow)

$tabs.TabPages.Add($windowsTab)
$tabs.TabPages.Add($officeTab)
$tabs.TabPages.Add($userTasksTab)
$tabs.TabPages.Add($newPcTab)

$splitContainer.Panel1.Controls.Add($tabs)

$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logPanel.BackColor = [System.Drawing.Color]::FromArgb(20, 22, 28)

$logToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$logToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
$logToolbar.Height = 48
$logToolbar.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$logToolbar.WrapContents = $false
$logToolbar.Padding = New-Object System.Windows.Forms.Padding(10, 9, 10, 8)
$logToolbar.BackColor = [System.Drawing.Color]::FromArgb(30, 34, 42)

$clearLogButton = New-Object System.Windows.Forms.Button
$clearLogButton.Text = "Clear Log"
$clearLogButton.Width = 100
$clearLogButton.Height = 28
$clearLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$clearLogButton.BackColor = [System.Drawing.Color]::FromArgb(84, 101, 255)
$clearLogButton.ForeColor = [System.Drawing.Color]::White
$clearLogButton.Add_Click({
    $script:logBox.Clear()
    Write-Log -Message "Log cleared." -Level "Info"
})

$openAppsButton = New-Object System.Windows.Forms.Button
$openAppsButton.Text = "Open C:\apps"
$openAppsButton.Width = 120
$openAppsButton.Height = 28
$openAppsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$openAppsButton.BackColor = [System.Drawing.Color]::FromArgb(56, 117, 215)
$openAppsButton.ForeColor = [System.Drawing.Color]::White
$openAppsButton.Add_Click({
    Invoke-UiAction -Name "Open Apps Folder" -Action {
        Ensure-Directory -Path $script:appsPath
        Start-LoggedProcess -FilePath "explorer.exe" -Arguments $script:appsPath -Description "C:\apps"
    }
})

$openOriginalButton = New-Object System.Windows.Forms.Button
$openOriginalButton.Text = "Open Original Script"
$openOriginalButton.Width = 150
$openOriginalButton.Height = 28
$openOriginalButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$openOriginalButton.BackColor = [System.Drawing.Color]::FromArgb(56, 117, 215)
$openOriginalButton.ForeColor = [System.Drawing.Color]::White
$openOriginalButton.Add_Click({
    Invoke-UiAction -Name "Open Original Script" -Action {
        Start-LoggedProcess -FilePath "notepad.exe" -Arguments $script:originalScriptPath -Description "Original script"
    }
})

$logToolbar.Controls.Add($clearLogButton)
$logToolbar.Controls.Add($openAppsButton)
$logToolbar.Controls.Add($openOriginalButton)

$script:logBox = New-Object System.Windows.Forms.RichTextBox
$script:logBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:logBox.ReadOnly = $true
$script:logBox.BackColor = [System.Drawing.Color]::FromArgb(17, 19, 24)
$script:logBox.ForeColor = [System.Drawing.Color]::Gainsboro
$script:logBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$script:logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None

$logPanel.Controls.Add($script:logBox)
$logPanel.Controls.Add($logToolbar)
$splitContainer.Panel2.Controls.Add($logPanel)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(238, 240, 244)
$script:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusLabel.Spring = $true
$script:statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$script:statusLabel.Text = "Ready"
$statusStrip.Items.Add($script:statusLabel) | Out-Null

$script:mainForm.Controls.Add($splitContainer)
$script:mainForm.Controls.Add($headerPanel)
$script:mainForm.Controls.Add($statusStrip)

$script:mainForm.Add_Shown({
    try {
        Set-Status -Message "Initializing support folders"
        Initialize-SupportFiles
        Set-Status -Message "Ready"
    } catch {
        Write-Log -Message "Initialization failed. $($_.Exception.Message)" -Level "Error"
        Set-Status -Message "Initialization failed"
    }
})

if ($SmokeTest) {
    try {
        Initialize-SupportFiles
        Set-Status -Message "Smoke test complete"
    } catch {
        Write-Log -Message "Smoke test failed. $($_.Exception.Message)" -Level "Error"
        throw
    } finally {
        $script:mainForm.Close()
        $script:mainForm.Dispose()
        [System.Windows.Forms.Application]::Exit()
    }
} else {
    [void]$script:mainForm.ShowDialog()
}
