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

# --- MODIFIED FUNCTION DEFINITIONS (Functions like Start-LongRunningJob, etc. remain unchanged from previous correct version) ---
# Placeholder for your other functions (Start-LongRunningJob, Change-PCName-Action, etc.)
# Ensure they are defined before New-ToolButton calls them or their scriptblock-returning counterparts.

# ... (All your other ...-Action and ...-ScriptBlock functions go here) ...
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

function Change-PCName-Action { Write-GuiLog "Mock Change-PCName-Action executed" -Color Green }
function Join-Domain-Action { param([string]$domainName, [string]$ouPath, [System.Management.Automation.PSCredential]$credential) Write-GuiLog "Mock Join-Domain-Action executed for $domainName" -Color Green }
function Repair-Windows-ScriptBlock { Write-GuiLog "Mock Repair-Windows-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Repair-Windows" } }
function Repair-SystemFiles-ScriptBlock { Write-GuiLog "Mock Repair-SystemFiles-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Repair-SystemFiles" } }
function Repair-Disk-Action { Write-GuiLog "Mock Repair-Disk-Action executed" -Color Green }
function Run-WindowsUpdateTroubleshooter-Action { Write-GuiLog "Mock Run-WindowsUpdateTroubleshooter-Action executed" -Color Green }
function Check-And-Repair-DISM-ScriptBlock { Write-GuiLog "Mock Check-And-Repair-DISM-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Check-And-Repair-DISM" } }
function Reset-Network-ScriptBlock { Write-GuiLog "Mock Reset-Network-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Reset-Network" } }
function Run-MemoryDiagnostic-Action { Write-GuiLog "Mock Run-MemoryDiagnostic-Action executed" -Color Green }
function Run-StartupRepair-Action { Write-GuiLog "Mock Run-StartupRepair-Action executed" -Color Green }
function Run-WindowsDefenderScan-ScriptBlock { Write-GuiLog "Mock Run-WindowsDefenderScan-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Run-WindowsDefenderScan" } }
function Reset-WindowsUpdateComponents-ScriptBlock { Write-GuiLog "Mock Reset-WindowsUpdateComponents-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Reset-WindowsUpdateComponents" } }
function Start-Teams-Action { Write-GuiLog "Mock Start-Teams-Action executed" -Color Green }
function Clear-TeamsCache-ScriptBlock { Write-GuiLog "Mock Clear-TeamsCache-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Clear-TeamsCache" } }
function List-InstalledApps-ScriptBlock { Write-GuiLog "Mock List-InstalledApps-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: List-InstalledApps" } }
function Network-Diagnostics-ScriptBlock { Write-GuiLog "Mock Network-Diagnostics-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Network-Diagnostics" } }
function Factory-Reset-Action { Write-GuiLog "Mock Factory-Reset-Action executed" -Color Green }
function Repair-Office-Action { Write-GuiLog "Mock Repair-Office-Action executed" -Color Green }
function Check-OfficeUpdates-Action { Write-GuiLog "Mock Check-OfficeUpdates-Action executed" -Color Green }
function Update-Windows-ScriptBlock { Write-GuiLog "Mock Update-Windows-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Update-Windows" } }
function Clean-TempFiles-ScriptBlock { Write-GuiLog "Mock Clean-TempFiles-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Clean-TempFiles" } }
function Map-Printer-Action { param([string]$PrinterConnectionName, [string]$PrinterFriendlyName) Write-GuiLog "Mock Map-Printer-Action for $PrinterFriendlyName ($PrinterConnectionName)" -Color Green }
function Install-AllPrinters-Action { Write-GuiLog "Mock Install-AllPrinters-Action executed" -Color Green }
function Open-NewPCFiles-Action { Write-GuiLog "Mock Open-NewPCFiles-Action executed" -Color Green }
function Download-And-Open-Ninite-ScriptBlock { Write-GuiLog "Mock Download-And-Open-Ninite-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Download-And-Open-Ninite" } }
function Download-MS-Teams-ScriptBlock { Write-GuiLog "Mock Download-MS-Teams-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Download-MS-Teams" } }
function Install-AdobeReader-ScriptBlock { Write-GuiLog "Mock Install-AdobeReader-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Install-AdobeReader" } }
function Remove-HPBloatware-ScriptBlock { Write-GuiLog "Mock Remove-HPBloatware-ScriptBlock called" -Color Green; return { Write-Host "Mock Job: Remove-HPBloatware" } }
# --- GUI Construction ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "RPI Repair & Setup Tool v1.4" # Updated version
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

# Function to create a button (REVISED WITH MORE DIAGNOSTICS)
function New-ToolButton {
    param(
        [string]$Text,
        $OnClickActionParam, # Renamed to avoid conflict if $OnClickAction is used elsewhere
        [System.Windows.Forms.Control]$ParentControl,
        [string]$JobScriptBlockFunctionNameParam, # Renamed
        [string]$OperationNameForJobParam # Renamed
    )

    # IMMEDIATE DIAGNOSTIC OF RECEIVED PARAMETERS
    Write-GuiLog "New-ToolButton RECEIVED PARAMS for '$Text':" -Color Magenta
    Write-GuiLog "  - OnClickActionParam is ScriptBlock: $($OnClickActionParam -is [scriptblock]) (Value: '$OnClickActionParam')" -Color Magenta -NoTimestamp
    Write-GuiLog "  - JobScriptBlockFunctionNameParam: '$JobScriptBlockFunctionNameParam'" -Color Magenta -NoTimestamp
    Write-GuiLog "  - OperationNameForJobParam: '$OperationNameForJobParam'" -Color Magenta -NoTimestamp

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.AutoSize = $true
    $button.MinimumSize = New-Object System.Drawing.Size(250, 0)
    $button.Padding = New-Object System.Windows.Forms.Padding(10,5,10,5)
    $button.Margin = New-Object System.Windows.Forms.Padding(5)

    # Store parameters intended for this button's action
    $buttonActionData = @{} # Initialize an empty hashtable for this button's specific data

    if ([!string]::IsNullOrWhiteSpace($JobScriptBlockFunctionNameParam) -and [!string]::IsNullOrWhiteSpace($OperationNameForJobParam)) {
        $buttonActionData.Type = "Job"
        $buttonActionData.JobFuncName = $JobScriptBlockFunctionNameParam
        $buttonActionData.OpName = $OperationNameForJobParam
        Write-GuiLog "  CONFIGURING AS JOB for '$Text'. JobFuncName: '$($buttonActionData.JobFuncName)', OpName: '$($buttonActionData.OpName)'" -Color DarkCyan
    } elseif ($OnClickActionParam -is [scriptblock]) {
        $buttonActionData.Type = "Direct"
        $buttonActionData.Action = $OnClickActionParam # Store the actual scriptblock
        Write-GuiLog "  CONFIGURING AS DIRECT for '$Text'. Action type: $($OnClickActionParam.GetType().FullName)" -Color DarkCyan
    } else {
        $button.Text = "$Text (Misconfigured - Check Logs)"
        $button.Enabled = $false
        Write-GuiLog "Button '$Text' is MISCONFIGURED: No valid job info OR direct action scriptblock provided based on initial checks." -Color Red
        Write-GuiLog "  - OnClickActionParam was: '$OnClickActionParam' (Is ScriptBlock: $($OnClickActionParam -is [scriptblock]))" -Color Red
        Write-GuiLog "  - JobScriptBlockFunctionNameParam was: '$JobScriptBlockFunctionNameParam'" -Color Red
        Write-GuiLog "  - OperationNameForJobParam was: '$OperationNameForJobParam'" -Color Red
        $ParentControl.Controls.Add($button)
        return $button
    }
    
    $button.Tag = $buttonActionData # Assign the prepared hashtable to Tag

    $button.Add_Click({
        param($sender, $eventArgs)
        $clickedButton = $sender -as [System.Windows.Forms.Button]
        $retrievedButtonData = $clickedButton.Tag

        Write-GuiLog "BUTTON CLICKED: '$($clickedButton.Text)'. Retrieved Tag Data:" -Color Blue
        $retrievedButtonData | Format-List | Out-String | ForEach-Object { Write-GuiLog $_ -Color Blue -NoTimestamp }

        if ($null -eq $retrievedButtonData -or $retrievedButtonData.Count -eq 0) {
            Write-GuiLog "CRITICAL ERROR (Button: $($clickedButton.Text)): Tag data is NULL or EMPTY!" -Color Red
            $statusLabel.Text = "Error: Button data missing for $($clickedButton.Text)"
            return
        }

        if ($retrievedButtonData.Type -eq "Job") {
            $statusLabel.Text = "Starting job: $($retrievedButtonData.OpName)..."
            Write-GuiLog "Preparing job '$($retrievedButtonData.OpName)'. ScriptBlock function from Tag: '$($retrievedButtonData.JobFuncName)'." -Color DarkCyan

            if ([string]::IsNullOrWhiteSpace($retrievedButtonData.JobFuncName)) {
                Write-GuiLog "CRITICAL ERROR (Button: $($clickedButton.Text)): JobFuncName in Tag is null or empty for operation '$($retrievedButtonData.OpName)'." -Color Red
                $statusLabel.Text = "Error preparing job: $($retrievedButtonData.OpName)"
                return
            }

            $scriptBlockFromFunc = $null
            try {
                $scriptBlockFromFunc = Invoke-Expression $retrievedButtonData.JobFuncName -ErrorAction Stop
            } catch {
                Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Failed to invoke '$($retrievedButtonData.JobFuncName)' for operation '$($retrievedButtonData.OpName)'. Error: $($_.Exception.Message)" -Color Red
                $statusLabel.Text = "Error invoking script function for: $($retrievedButtonData.OpName)"
                return
            }
            
            if (-not $scriptBlockFromFunc) {
                Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Function '$($retrievedButtonData.JobFuncName)' did NOT return a scriptblock (returned null) for operation '$($retrievedButtonData.OpName)'." -Color Red
                $statusLabel.Text = "Error: Script function returned null for: $($retrievedButtonData.OpName)"
                return
            }
            if ($scriptBlockFromFunc -isnot [scriptblock]) {
                 Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Function '$($retrievedButtonData.JobFuncName)' returned a value of type '$($scriptBlockFromFunc.GetType().FullName)' instead of [scriptblock] for operation '$($retrievedButtonData.OpName)'." -Color Red
                 $statusLabel.Text = "Error: Script function returned wrong type for: $($retrievedButtonData.OpName)"
                 return
            }
            Start-LongRunningJob -ScriptBlock $scriptBlockFromFunc -OperationName $retrievedButtonData.OpName -ButtonToDisable $clickedButton
        
        } elseif ($retrievedButtonData.Type -eq "Direct") {
            $statusLabel.Text = "Executing: $($clickedButton.Text)..."
            try {
                if ($retrievedButtonData.Action -is [scriptblock]) {
                    & $retrievedButtonData.Action # Direct invocation
                    Write-GuiLog "Direct action '$($clickedButton.Text)' completed." -Color Green
                } else {
                    Write-GuiLog "ERROR (Button: $($clickedButton.Text)): Action in Tag is not a scriptblock. Type: $($retrievedButtonData.Action.GetType().FullName)" -Color Red
                    Write-GuiLog "  Value of Action in Tag: '$($retrievedButtonData.Action)'" -Color Red
                }
            } catch {
                Write-GuiLog "Error during direct action '$($clickedButton.Text)': $($_.Exception.Message)" -Color Red
                if ($retrievedButtonData.Action -is [scriptblock]) {
                    Write-GuiLog "Failing scriptblock for direct action was: $($retrievedButtonData.Action.ToString())" -Color DarkRed
                }
            }
            $statusLabel.Text = "Finished: $($clickedButton.Text). Check log."
        } else {
             Write-GuiLog "CRITICAL ERROR (Button: $($clickedButton.Text)): Unknown action type in Tag: '$($retrievedButtonData.Type)'" -Color Red
             $statusLabel.Text = "Error: Unknown button action type for $($clickedButton.Text)"
        }
    })
    
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
New-ToolButton -Text "1: Open New PC Files Folder" -OnClickActionParam {Open-NewPCFiles-Action} -ParentControl $panelNewPC
New-ToolButton -Text "2: Download & Run Ninite" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Download-And-Open-Ninite-ScriptBlock" -OperationNameForJobParam "Ninite Download & Run"
New-ToolButton -Text "3: Download MS Teams (New)" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Download-MS-Teams-ScriptBlock" -OperationNameForJobParam "MS Teams Download"
New-ToolButton -Text "4: Change PC Name (Restarts PC)" -OnClickActionParam {Change-PCName-Action} -ParentControl $panelNewPC

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

New-ToolButton -Text "6: Update Windows (May Restart)" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Update-Windows-ScriptBlock" -OperationNameForJobParam "Windows Update"
New-ToolButton -Text "7: Install Adobe Acrobat Reader" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Install-AdobeReader-ScriptBlock" -OperationNameForJobParam "Adobe Reader Install"
New-ToolButton -Text "8: Remove HP Bloatware" -ParentControl $panelNewPC -JobScriptBlockFunctionNameParam "Remove-HPBloatware-ScriptBlock" -OperationNameForJobParam "HP Bloatware Removal"
$tabControl.Controls.Add($tabNewPC)

# == Windows Repairs Tab ==
$tabWindows = New-Object System.Windows.Forms.TabPage; $tabWindows.Text = "Windows Repairs"
$panelWindows = New-ButtonFlowPanel; $tabWindows.Controls.Add($panelWindows)
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
$tabOffice = New-Object System.Windows.Forms.TabPage; $tabOffice.Text = "Office Repairs"
$panelOffice = New-ButtonFlowPanel; $tabOffice.Controls.Add($panelOffice)
New-ToolButton -Text "1: Repair Microsoft Office" -OnClickActionParam {Repair-Office-Action} -ParentControl $panelOffice
New-ToolButton -Text "2: Check for Microsoft Office Updates" -OnClickActionParam {Check-OfficeUpdates-Action} -ParentControl $panelOffice
$tabControl.Controls.Add($tabOffice)

# == User Tasks Tab ==
$tabUserTasks = New-Object System.Windows.Forms.TabPage; $tabUserTasks.Text = "User Tasks"
$panelUserTasks = New-ButtonFlowPanel; $tabUserTasks.Controls.Add($panelUserTasks)
New-ToolButton -Text "1: Clean Temp Files" -ParentControl $panelUserTasks -JobScriptBlockFunctionNameParam "Clean-TempFiles-ScriptBlock" -OperationNameForJobParam "Clean Temp Files"

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
    New-ToolButton -Text "Map $($entry.Key)" -OnClickActionParam $actionScriptBlock -ParentControl $flowPrinters
}
New-ToolButton -Text "Install All Printers (via VBS)" -OnClickActionParam {Install-AllPrinters-Action} -ParentControl $flowPrinters
$groupPrinters.Controls.Add($flowPrinters); $panelUserTasks.Controls.Add($groupPrinters)

New-ToolButton -Text "3: Clear Teams Cache & Restart Teams" -ParentControl $panelUserTasks -JobScriptBlockFunctionNameParam "Clear-TeamsCache-ScriptBlock" -OperationNameForJobParam "Clear Teams Cache"
New-ToolButton -Text "BONUS: Start Microsoft Teams" -OnClickActionParam {Start-Teams-Action} -ParentControl $panelUserTasks
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
