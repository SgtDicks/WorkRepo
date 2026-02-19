<#
Fix-1355U-BatteryFreeze.ps1

Targets common battery-freeze causes on 13th-gen Intel U (incl i7-1355U):
- Sets DC (battery) Processor Max = 99% and Min = 5% on the active power plan
- Optionally disables Intel Dynamic Tuning / DPTF services and devices
- Optionally opens a pre-addressed Outlook email with the log attached (user chooses to send)

Run as Administrator.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$DisableIntelDynamicTuning = $true,
    [switch]$SetPowerPlanProcessorCaps = $true,

    # Outlook draft email options
    [switch]$OpenOutlookDraft = $true,
    [string]$EmailTo = "Aaron.Bycroft@rpinfrastructure.com.au",

    # If not provided, subject will be: 1355U config change for <username>
    [string]$EmailSubject = "",

    [string]$LogPath = "$env:ProgramData\BatteryFreezeFix\Fix-1355U-BatteryFreeze.log"
)

# ------------------------ Helpers ------------------------

function Ensure-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = [Security.Principal.WindowsPrincipal]::new($id)
    if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Please run PowerShell as Administrator."
    }
}

function Write-Log {
    param([Parameter(Mandatory)][string]$Message)

    try {
        $dir = Split-Path -Parent $LogPath
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $line  = "[$stamp] $Message"

        $line | Tee-Object -FilePath $LogPath -Append | Out-Null
    }
    catch {
        Write-Warning "Logging failed: $($_.Exception.Message)"
        Write-Warning "Original log message: $Message"
    }
}

function Get-RunAsUserForSubject {
    # Prefer DOMAIN\User from Win32_ComputerSystem, fallback to env var
    $u = $null
    try { $u = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName } catch {}
    if (-not $u) { $u = $env:USERNAME }
    return $u
}

function Get-ActiveSchemeGuid {
    $out = powercfg /GETACTIVESCHEME 2>$null
    if ($out -match '([0-9a-fA-F-]{36})') { return $Matches[1] }
    throw "Could not determine active power scheme."
}

function Invoke-PowerCfg {
    param(
        [Parameter(Mandatory)][string]$Arguments,
        [string]$What = "powercfg $Arguments"
    )

    if ($PSCmdlet.ShouldProcess($What, "Execute")) {
        Write-Log "Running: powercfg $Arguments"
        $p = Start-Process -FilePath "powercfg.exe" -ArgumentList $Arguments -NoNewWindow -Wait -PassThru -ErrorAction Stop
        if ($p.ExitCode -ne 0) {
            throw "powercfg.exe exited with code $($p.ExitCode) while running: $Arguments"
        }
    }
}

function Set-PowerCfgValue {
    param(
        [Parameter(Mandatory)][string]$SchemeGuid,
        [Parameter(Mandatory)][string]$SubGroupGuid,
        [Parameter(Mandatory)][string]$SettingGuid,
        [Parameter(Mandatory)][ValidateSet("AC","DC")][string]$Mode,
        [Parameter(Mandatory)][ValidateRange(0,100)][int]$Percent
    )

    $args = if ($Mode -eq "AC") {
        "/SETACVALUEINDEX $SchemeGuid $SubGroupGuid $SettingGuid $Percent"
    } else {
        "/SETDCVALUEINDEX $SchemeGuid $SubGroupGuid $SettingGuid $Percent"
    }

    if ($PSCmdlet.ShouldProcess("$Mode Processor setting $SettingGuid", "Set to $Percent%")) {
        Invoke-PowerCfg -Arguments $args -What "$Mode value index $SettingGuid = $Percent"
    }
}

function Apply-ProcessorCaps {
    # Powercfg GUIDs
    $SUB_PROCESSOR   = "54533251-82be-4824-96c1-47b60b740d00"
    $PROCTHROTTLEMIN = "893dee8e-2bef-41e0-89c6-b55d0929964c"
    $PROCTHROTTLEMAX = "bc5038f7-23e0-4960-96da-33abaf5935ec"

    $scheme = Get-ActiveSchemeGuid
    Write-Log "Active power scheme: $scheme"

    # AC: Min 5, Max 100 (normal plugged-in behavior)
    Set-PowerCfgValue -SchemeGuid $scheme -SubGroupGuid $SUB_PROCESSOR -SettingGuid $PROCTHROTTLEMIN -Mode AC -Percent 5
    Set-PowerCfgValue -SchemeGuid $scheme -SubGroupGuid $SUB_PROCESSOR -SettingGuid $PROCTHROTTLEMAX -Mode AC -Percent 100

    # DC (Battery): Min 5, Max 99 (disables Turbo on battery; common stability fix)
    Set-PowerCfgValue -SchemeGuid $scheme -SubGroupGuid $SUB_PROCESSOR -SettingGuid $PROCTHROTTLEMIN -Mode DC -Percent 5
    Set-PowerCfgValue -SchemeGuid $scheme -SubGroupGuid $SUB_PROCESSOR -SettingGuid $PROCTHROTTLEMAX -Mode DC -Percent 99

    if ($PSCmdlet.ShouldProcess("Power scheme $scheme", "Activate changes")) {
        Write-Log "Activating scheme to apply changes: $scheme"
        Invoke-PowerCfg -Arguments "/SETACTIVE $scheme" -What "Set active scheme $scheme"
    }

    Write-Log "Processor caps applied (AC max 100 / DC max 99; min 5 both)."
}

function Disable-IntelTuningServices {
    # Match common display names and service names (varies by OEM)
    $patterns = @(
        "*Dynamic Tuning*",
        "*Dynamic Platform*",
        "*DPTF*",
        "*Intel(R) Innovation Platform Framework*",
        "*Intel*Platform*Thermal*",
        "*Intel*DTT*"
    )

    $services = Get-Service | Where-Object {
        $dn = $_.DisplayName
        $sn = $_.Name
        ($patterns | Where-Object { $dn -like $_ -or $sn -like $_ } | Select-Object -First 1) -ne $null
    } | Sort-Object -Property DisplayName -Unique

    if (-not $services) {
        Write-Log "No matching Intel tuning services found."
        return
    }

    foreach ($svc in $services) {
        try {
            $cim = Get-CimInstance Win32_Service -Filter "Name='$($svc.Name)'" -ErrorAction SilentlyContinue
            $startMode = if ($cim) { $cim.StartMode } else { "Unknown" }

            Write-Log "Found service: $($svc.DisplayName) [$($svc.Name)] Status=$($svc.Status) StartType=$startMode"

            if ($PSCmdlet.ShouldProcess("Service $($svc.Name)", "Stop + Disable")) {

                if ($svc.Status -ne "Stopped") {
                    Write-Log "Stopping service: $($svc.Name)"
                    Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
                }

                Write-Log "Disabling service: $($svc.Name)"
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log "Failed to change service $($svc.Name): $($_.Exception.Message)"
        }
    }
}

function Disable-IntelTuningDevices {
    if (-not (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue)) {
        Write-Log "PnPDevice cmdlets not available; skipping device disable."
        return
    }

    $namePatterns = @(
        "*Intel*Dynamic Tuning*",
        "*Intel*Dynamic Platform*Thermal*",
        "*Intel*DPTF*",
        "*Dynamic Tuning*",
        "*DPTF*"
    )

    $devices = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue | Where-Object {
        $fn = $_.FriendlyName
        if (-not $fn) { return $false }
        ($namePatterns | Where-Object { $fn -like $_ } | Select-Object -First 1) -ne $null
    }

    if (-not $devices) {
        Write-Log "No matching Intel tuning devices found."
        return
    }

    foreach ($dev in $devices) {
        try {
            Write-Log "Found device: $($dev.FriendlyName) [$($dev.InstanceId)] Status=$($dev.Status)"

            if ($PSCmdlet.ShouldProcess("Device $($dev.FriendlyName)", "Disable")) {
                Disable-PnpDevice -InstanceId $dev.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                Write-Log "Disabled device: $($dev.FriendlyName)"
            }
        }
        catch {
            Write-Log "Failed to disable device $($dev.FriendlyName): $($_.Exception.Message)"
        }
    }
}

function New-OutlookDraftWithAttachment {
    param(
        [Parameter(Mandatory)][string]$ToAddress,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Body,
        [Parameter(Mandatory)][string]$AttachmentPath
    )

    if (-not (Test-Path $AttachmentPath)) {
        Write-Log "Outlook draft not created: attachment not found at $AttachmentPath"
        return
    }

    try {
        # Try bind to running Outlook first
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        }
        catch {
            $outlook = New-Object -ComObject Outlook.Application
        }

        $mail = $outlook.CreateItem(0) # 0 = olMailItem
        $mail.To      = $ToAddress
        $mail.Subject = $Subject
        $mail.Body    = $Body
        [void]$mail.Attachments.Add($AttachmentPath)

        # Display (do NOT send). Modal inspector keeps it frontmost.
        $mail.Display($true) | Out-Null

        Write-Log "Opened Outlook compose window with log attached (user can send or discard)."
    }
    catch {
        Write-Log "Failed to open Outlook draft: $($_.Exception.Message)"
    }
}

# ------------------------ Main ------------------------

try {
    Ensure-Admin
    Write-Log "=== Starting battery freeze fix ==="

    if ($SetPowerPlanProcessorCaps) {
        Write-Log "Applying power plan processor caps..."
        Apply-ProcessorCaps
    }
    else {
        Write-Log "Skipping power plan changes (SetPowerPlanProcessorCaps=$SetPowerPlanProcessorCaps)."
    }

    if ($DisableIntelDynamicTuning) {
        Write-Log "Disabling Intel Dynamic Tuning / DPTF services and devices..."
        Disable-IntelTuningServices
        Disable-IntelTuningDevices
    }
    else {
        Write-Log "Skipping Intel Dynamic Tuning/DPTF disables (DisableIntelDynamicTuning=$DisableIntelDynamicTuning)."
    }

    Write-Log "=== Completed battery freeze fix. Reboot recommended. ==="

    if ($OpenOutlookDraft) {
        $runAsUser = Get-RunAsUserForSubject

        if ([string]::IsNullOrWhiteSpace($EmailSubject)) {
            $EmailSubject = "1355U config change for $runAsUser"
        }

        $body = @"
Hi Aaron,

Attached is the log from the 1355U battery freeze configuration change run for:

$runAsUser

Reboot is recommended to fully apply driver/service/device state changes.

Thanks,
"@

        Write-Log "Opening Outlook draft email to $EmailTo with subject: $EmailSubject"
        New-OutlookDraftWithAttachment -ToAddress $EmailTo -Subject $EmailSubject -Body $body -AttachmentPath $LogPath
    }
    else {
        Write-Log "Skipping Outlook draft creation (OpenOutlookDraft=$OpenOutlookDraft)."
    }

    Write-Host "`nDone. Log written to: $LogPath`nReboot is recommended to fully apply driver/service/device state changes.`n"
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    throw
}
