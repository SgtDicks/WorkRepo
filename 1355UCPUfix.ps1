<# 
Fix-1355U-BatteryFreeze.ps1
Targets common battery-freeze causes on 13th-gen Intel U (incl i7-1355U):
- Sets DC (battery) Processor Max = 99% and Min = 5% on the active power plan
- Optionally disables Intel Dynamic Tuning / DPTF services and devices

Run as Administrator.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$DisableIntelDynamicTuning = $true,
    [switch]$SetPowerPlanProcessorCaps = $true,
    [string]$LogPath = "$env:ProgramData\BatteryFreezeFix\Fix-1355U-BatteryFreeze.log"
)

# ------------------------ Helpers ------------------------

function Ensure-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Please run PowerShell as Administrator."
    }
}

function Write-Log {
    param([string]$Message)
    $dir = Split-Path -Parent $LogPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$stamp] $Message"
    $line | Tee-Object -FilePath $LogPath -Append
}

function Get-ActiveSchemeGuid {
    $out = powercfg /GETACTIVESCHEME 2>$null
    # Example: "Power Scheme GUID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  (Balanced)"
    if ($out -match '([0-9a-fA-F-]{36})') { return $Matches[1] }
    throw "Could not determine active power scheme."
}

function Set-PowerCfgValue {
    param(
        [Parameter(Mandatory)][string]$SchemeGuid,
        [Parameter(Mandatory)][string]$SubGroupGuid,
        [Parameter(Mandatory)][string]$SettingGuid,
        [Parameter(Mandatory)][ValidateSet("AC","DC")][string]$Mode,
        [Parameter(Mandatory)][int]$Percent
    )

    $cmd = if ($Mode -eq "AC") {
        "powercfg /SETACVALUEINDEX $SchemeGuid $SubGroupGuid $SettingGuid $Percent"
    } else {
        "powercfg /SETDCVALUEINDEX $SchemeGuid $SubGroupGuid $SettingGuid $Percent"
    }

    if ($PSCmdlet.ShouldProcess("$Mode Processor setting $SettingGuid", "Set to $Percent%")) {
        Write-Log "Running: $cmd"
        cmd.exe /c $cmd | Out-Null
    }
}

function Apply-ProcessorCaps {
    # Powercfg GUIDs
    $SUB_PROCESSOR      = "54533251-82be-4824-96c1-47b60b740d00"
    $PROCTHROTTLEMIN    = "893dee8e-2bef-41e0-89c6-b55d0929964c"
    $PROCTHROTTLEMAX    = "bc5038f7-23e0-4960-96da-33abaf5935ec"

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
        powercfg /SETACTIVE $scheme | Out-Null
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
        $patterns | ForEach-Object { if ($dn -like $_ -or $sn -like $_) { return $true } }
        return $false
    } | Sort-Object DisplayName -Unique

    if (-not $services) {
        Write-Log "No matching Intel tuning services found."
        return
    }

    foreach ($svc in $services) {
        try {
            Write-Log "Found service: $($svc.DisplayName) [$($svc.Name)] Status=$($svc.Status) StartType=$((Get-CimInstance Win32_Service -Filter "Name='$($svc.Name)'" ).StartMode)"

            if ($PSCmdlet.ShouldProcess("Service $($svc.Name)", "Stop + Disable")) {
                if ($svc.Status -ne "Stopped") {
                    Write-Log "Stopping service: $($svc.Name)"
                    Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
                }
                Write-Log "Disabling service: $($svc.Name)"
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Log "Failed to change service $($svc.Name): $($_.Exception.Message)"
        }
    }
}

function Disable-IntelTuningDevices {
    # Requires PnPDevice cmdlets (present on Win10/11)
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
        foreach ($p in $namePatterns) {
            if ($fn -like $p) { return $true }
        }
        return $false
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
        } catch {
            Write-Log "Failed to disable device $($dev.FriendlyName): $($_.Exception.Message)"
        }
    }
}

# ------------------------ Main ------------------------

try {
    Ensure-Admin
    Write-Log "=== Starting battery freeze fix ==="

    if ($SetPowerPlanProcessorCaps) {
        Write-Log "Applying power plan processor caps..."
        Apply-ProcessorCaps
    } else {
        Write-Log "Skipping power plan changes (SetPowerPlanProcessorCaps=$SetPowerPlanProcessorCaps)."
    }

    if ($DisableIntelDynamicTuning) {
        Write-Log "Disabling Intel Dynamic Tuning / DPTF services and devices..."
        Disable-IntelTuningServices
        Disable-IntelTuningDevices
    } else {
        Write-Log "Skipping Intel Dynamic Tuning/DPTF disables (DisableIntelDynamicTuning=$DisableIntelDynamicTuning)."
    }

    Write-Log "=== Completed battery freeze fix. Reboot recommended. ==="
    Write-Host "`nDone. Log written to: $LogPath`nReboot is recommended to fully apply driver/service/device state changes.`n"
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    throw
}
