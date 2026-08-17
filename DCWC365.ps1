<#
.SYNOPSIS
    Microsoft 365 Tenant Migration - Login Reset GUI
 
.DESCRIPTION
    Clears cached Microsoft 365 / Office authentication information
    after a tenant-to-tenant migration.
 
    - No reboot
    - Does NOT run dsregcmd /leave
    - Does NOT rejoin Entra ID
    - Clears Office identity/token caches
    - Clears relevant Windows Credential Manager entries
    - Clears Office activation cache
    - Provides GUI logging
 
.NOTES
    Run in the affected user's Windows session.
#>
 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 
$ErrorActionPreference = "SilentlyContinue"
 
 
# ============================================================
# GUI
# ============================================================
 
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft 365 Tenant Login Reset"
$form.Size = New-Object System.Drawing.Size(720, 620)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
 
 
# ------------------------------------------------------------
# Header
# ------------------------------------------------------------
 
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Microsoft 365 Tenant Migration"
$lblTitle.Location = New-Object System.Drawing.Point(25, 20)
$lblTitle.Size = New-Object System.Drawing.Size(650, 35)
$lblTitle.Font = New-Object System.Drawing.Font(
    "Segoe UI",
    18,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($lblTitle)
 
 
$lblDescription = New-Object System.Windows.Forms.Label
$lblDescription.Text = "Clear old Microsoft 365 authentication information and prepare Office to sign into the new tenant."
$lblDescription.Location = New-Object System.Drawing.Point(28, 60)
$lblDescription.Size = New-Object System.Drawing.Size(650, 40)
$form.Controls.Add($lblDescription)
 
 
# ------------------------------------------------------------
# Current Windows User
# ------------------------------------------------------------
 
$lblCurrentUserTitle = New-Object System.Windows.Forms.Label
$lblCurrentUserTitle.Text = "Windows user:"
$lblCurrentUserTitle.Location = New-Object System.Drawing.Point(30, 110)
$lblCurrentUserTitle.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblCurrentUserTitle)
 
 
$lblCurrentUser = New-Object System.Windows.Forms.Label
$lblCurrentUser.Text = "$env:USERDOMAIN\$env:USERNAME"
$lblCurrentUser.Location = New-Object System.Drawing.Point(150, 110)
$lblCurrentUser.Size = New-Object System.Drawing.Size(500, 20)
$lblCurrentUser.Font = New-Object System.Drawing.Font(
    "Segoe UI",
    9,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($lblCurrentUser)
 
 
# ------------------------------------------------------------
# New Microsoft 365 UPN
# ------------------------------------------------------------
 
$lblUPN = New-Object System.Windows.Forms.Label
$lblUPN.Text = "New Microsoft 365 username:"
$lblUPN.Location = New-Object System.Drawing.Point(30, 145)
$lblUPN.Size = New-Object System.Drawing.Size(200, 22)
$form.Controls.Add($lblUPN)
 
 
$txtUPN = New-Object System.Windows.Forms.TextBox
$txtUPN.Location = New-Object System.Drawing.Point(30, 170)
$txtUPN.Size = New-Object System.Drawing.Size(640, 28)
$txtUPN.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($txtUPN)
 
 
# ------------------------------------------------------------
# Reset Button
# ------------------------------------------------------------
 
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "RESET MICROSOFT 365 LOGIN"
$btnReset.Location = New-Object System.Drawing.Point(30, 215)
$btnReset.Size = New-Object System.Drawing.Size(310, 45)
$btnReset.Font = New-Object System.Drawing.Font(
    "Segoe UI",
    10,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($btnReset)
 
 
# ------------------------------------------------------------
# Open Word Button
# ------------------------------------------------------------
 
$btnWord = New-Object System.Windows.Forms.Button
$btnWord.Text = "Open Word"
$btnWord.Location = New-Object System.Drawing.Point(360, 215)
$btnWord.Size = New-Object System.Drawing.Size(150, 45)
$btnWord.Enabled = $false
$form.Controls.Add($btnWord)
 
 
# ------------------------------------------------------------
# Open Excel Button
# ------------------------------------------------------------
 
$btnExcel = New-Object System.Windows.Forms.Button
$btnExcel.Text = "Open Excel"
$btnExcel.Location = New-Object System.Drawing.Point(520, 215)
$btnExcel.Size = New-Object System.Drawing.Size(150, 45)
$btnExcel.Enabled = $false
$form.Controls.Add($btnExcel)
 
 
# ------------------------------------------------------------
# Progress Bar
# ------------------------------------------------------------
 
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(30, 280)
$progress.Size = New-Object System.Drawing.Size(640, 20)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)
 
 
# ------------------------------------------------------------
# Status
# ------------------------------------------------------------
 
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(30, 310)
$lblStatus.Size = New-Object System.Drawing.Size(640, 25)
$lblStatus.Font = New-Object System.Drawing.Font(
    "Segoe UI",
    9,
    [System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($lblStatus)
 
 
# ------------------------------------------------------------
# Log Window
# ------------------------------------------------------------
 
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(30, 345)
$txtLog.Size = New-Object System.Drawing.Size(640, 180)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)
 
 
# ------------------------------------------------------------
# Close Button
# ------------------------------------------------------------
 
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(550, 535)
$btnClose.Size = New-Object System.Drawing.Size(120, 32)
$form.Controls.Add($btnClose)
 
 
# ============================================================
# Functions
# ============================================================
 
function Write-GUILog {
 
    param (
        [string]$Message
    )
 
    $timestamp = Get-Date -Format "HH:mm:ss"
 
    $txtLog.AppendText(
        "[$timestamp] $Message`r`n"
    )
 
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
 
    [System.Windows.Forms.Application]::DoEvents()
}
 
 
function Set-Progress {
 
    param (
        [int]$Value,
        [string]$Status
    )
 
    $progress.Value = $Value
    $lblStatus.Text = $Status
 
    [System.Windows.Forms.Application]::DoEvents()
}
 
 
# ============================================================
# Main Reset Function
# ============================================================
 
function Reset-M365Login {
 
    $btnReset.Enabled = $false
    $btnWord.Enabled = $false
    $btnExcel.Enabled = $false
 
    $txtLog.Clear()
 
    try {
 
        # ====================================================
        # STEP 1
        # Close Microsoft applications
        # ====================================================
 
        Set-Progress 10 "Closing Microsoft 365 applications..."
 
        Write-GUILog "Closing Microsoft 365 applications."
 
        $Processes = @(
            "OUTLOOK",
            "WINWORD",
            "EXCEL",
            "POWERPNT",
            "MSACCESS",
            "ONENOTE",
            "VISIO",
            "MSPUB",
            "Teams",
            "ms-teams",
            "OneDrive"
        )
 
        foreach ($Process in $Processes) {
 
            $RunningProcess = Get-Process `
                -Name $Process `
                -ErrorAction SilentlyContinue
 
            if ($RunningProcess) {
 
                Write-GUILog "Closing $Process"
 
                $RunningProcess |
                    Stop-Process `
                        -Force `
                        -ErrorAction SilentlyContinue
            }
        }
 
        Start-Sleep -Seconds 1
 
 
        # ====================================================
        # STEP 2
        # Credential Manager
        # ====================================================
 
        Set-Progress 25 "Removing cached Office credentials..."
 
        Write-GUILog "Checking Windows Credential Manager."
 
        $CredentialOutput = cmdkey.exe /list
 
        $Targets = foreach ($Line in $CredentialOutput) {
 
            if ($Line -match "Target:\s*(.+)$") {
 
                $Target = $Matches[1].Trim()
 
                if (
                    $Target -match "MicrosoftOffice" -or
                    $Target -match "Office16" -or
                    $Target -match "ADAL" -or
                    $Target -match "OneAuth"
                ) {
 
                    $Target
                }
            }
        }
 
 
        if ($Targets) {
 
            foreach ($Target in $Targets) {
 
                Write-GUILog "Removing credential: $Target"
 
                cmdkey.exe /delete:$Target |
                    Out-Null
            }
 
        }
        else {
 
            Write-GUILog "No matching Credential Manager entries found."
        }
 
 
        # ====================================================
        # STEP 3
        # Office Identity
        # ====================================================
 
        Set-Progress 40 "Clearing Office identities..."
 
        Write-GUILog "Clearing Office identity registry information."
 
        $OfficeIdentityPath =
            "HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities"
 
 
        if (Test-Path $OfficeIdentityPath) {
 
            Remove-Item `
                $OfficeIdentityPath `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
 
            Write-GUILog "Office identities cleared."
 
        }
        else {
 
            Write-GUILog "No Office identity registry entries found."
        }
 
 
        # ====================================================
        # STEP 4
        # Roaming Office identities
        # ====================================================
 
        Set-Progress 50 "Clearing Office roaming identities..."
 
        $RoamingIdentityPath =
            "HKCU:\Software\Microsoft\Office\16.0\Common\Roaming\Identities"
 
 
        if (Test-Path $RoamingIdentityPath) {
 
            Remove-Item `
                $RoamingIdentityPath `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
 
            Write-GUILog "Office roaming identities cleared."
 
        }
        else {
 
            Write-GUILog "No roaming Office identities found."
        }
 
 
        # ====================================================
        # STEP 5
        # OneAuth
        # ====================================================
 
        Set-Progress 60 "Clearing Microsoft OneAuth cache..."
 
        $OneAuth =
            "$env:LOCALAPPDATA\Microsoft\OneAuth"
 
 
        if (Test-Path $OneAuth) {
 
            Write-GUILog "Clearing OneAuth cache."
 
            Remove-Item `
                "$OneAuth\*" `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
 
        }
        else {
 
            Write-GUILog "OneAuth cache not present."
        }
 
 
        # ====================================================
        # STEP 6
        # IdentityCache
        # ====================================================
 
        Set-Progress 70 "Clearing Microsoft IdentityCache..."
 
        $IdentityCache =
            "$env:LOCALAPPDATA\Microsoft\IdentityCache"
 
 
        if (Test-Path $IdentityCache) {
 
            Write-GUILog "Clearing IdentityCache."
 
            Remove-Item `
                "$IdentityCache\*" `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
 
        }
        else {
 
            Write-GUILog "IdentityCache not present."
        }
 
 
        # ====================================================
        # STEP 7
        # Office licensing cache
        # ====================================================
 
        Set-Progress 80 "Clearing Microsoft 365 activation cache..."
 
        $LicenceFolder =
            "$env:LOCALAPPDATA\Microsoft\Office\Licenses"
 
 
        if (Test-Path $LicenceFolder) {
 
            Write-GUILog "Clearing Office licence cache."
 
            Remove-Item `
                "$LicenceFolder\*" `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue
 
        }
        else {
 
            Write-GUILog "Office licence cache not present."
        }
 
 
        # ====================================================
        # STEP 8
        # Stop authentication processes
        # ====================================================
 
        Set-Progress 90 "Refreshing authentication processes..."
 
        Write-GUILog "Refreshing Windows authentication processes."
 
        $AuthProcesses = @(
            "Microsoft.AAD.BrokerPlugin",
            "TokenBroker"
        )
 
 
        foreach ($Process in $AuthProcesses) {
 
            Get-Process `
                -Name $Process `
                -ErrorAction SilentlyContinue |
                Stop-Process `
                    -Force `
                    -ErrorAction SilentlyContinue
        }
 
 
        Start-Sleep -Seconds 1
 
 
        # ====================================================
        # COMPLETE
        # ====================================================
 
        Set-Progress 100 "Microsoft 365 login reset complete"
 
        Write-GUILog ""
        Write-GUILog "======================================"
        Write-GUILog "Microsoft 365 login reset COMPLETE."
        Write-GUILog "NO REBOOT REQUIRED."
        Write-GUILog "======================================"
 
 
        if (
            -not [string]::IsNullOrWhiteSpace(
                $txtUPN.Text
            )
        ) {
 
            Write-GUILog ""
            Write-GUILog "New account:"
            Write-GUILog $txtUPN.Text
        }
 
 
        Write-GUILog ""
        Write-GUILog "Open Word or Excel and sign into the NEW tenant."
 
 
        $lblStatus.Text =
            "Complete - sign into Microsoft 365 with the NEW tenant account."
 
        $btnWord.Enabled = $true
        $btnExcel.Enabled = $true
 
 
        [System.Windows.Forms.MessageBox]::Show(
            "Microsoft 365 cached login information has been cleared.`r`n`r`nNo reboot is required.`r`n`r`nOpen Word or Excel and sign in using the NEW tenant account.",
            "Microsoft 365 Reset Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
 
    }
    catch {
 
        Write-GUILog ""
        Write-GUILog "ERROR: $($_.Exception.Message)"
 
        $lblStatus.Text = "Reset encountered an error."
 
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred:`r`n`r`n$($_.Exception.Message)",
            "Microsoft 365 Reset Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
 
    finally {
 
        $btnReset.Enabled = $true
    }
}
 
 
# ============================================================
# Button Events
# ============================================================
 
$btnReset.Add_Click({
 
    $Result = [System.Windows.Forms.MessageBox]::Show(
        "This will close Outlook, Teams, OneDrive and all open Office applications.`r`n`r`nAny unsaved Office documents could be lost.`r`n`r`nContinue?",
        "Confirm Microsoft 365 Reset",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
 
    if (
        $Result -eq
        [System.Windows.Forms.DialogResult]::Yes
    ) {
 
        Reset-M365Login
    }
})
 
 
$btnWord.Add_Click({
 
    Write-GUILog "Opening Microsoft Word."
 
    try {
 
        Start-Process "winword.exe"
 
    }
    catch {
 
        [System.Windows.Forms.MessageBox]::Show(
            "Microsoft Word could not be started.",
            "Word",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})
 
 
$btnExcel.Add_Click({
 
    Write-GUILog "Opening Microsoft Excel."
 
    try {
 
        Start-Process "excel.exe"
 
    }
    catch {
 
        [System.Windows.Forms.MessageBox]::Show(
            "Microsoft Excel could not be started.",
            "Excel",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})
 
 
$btnClose.Add_Click({
 
    $form.Close()
 
})
 
 
# ============================================================
# Start GUI
# ============================================================
 
[void]$form.ShowDialog()
