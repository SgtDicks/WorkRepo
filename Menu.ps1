# Array of script URLs
$scripts = @(
    @{
        Name = "Folder Inheritance Check"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/Folder%20Inheritance%20Check.ps1"
    },
    @{
        Name = "Setup Machine"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/SetupMachine.ps1"
    },
    @{
        Name = "Watch for External"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/Watch%20for%20external.ps1"
    },
    @{
        Name = "Find and Remove Meeting"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/find%20and%20remove%20meeting.ps1"
    },
    @{
        Name = "Find Email Address in Exchange DL"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/find_emailaddress_in_exchange_DL.PS1"
    },
    @{
        Name = "Folder Permissions Export"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/folder%20permissions%20export.ps1"
    },
    @{
        Name = "Latest YouTube Downloader"
        Url  = "https://raw.githubusercontent.com/SgtDicks/WorkRepo/main/latest%20youtube%20downloader.ps1"
    }
)

# Function to display menu and run selected script
function Show-ScriptMenu {
    Write-Host "Select a script to run:" -ForegroundColor Green
    $counter = 1
    foreach ($script in $scripts) {
        Write-Host "$counter. $($script.Name)"
        $counter++
    }

    $selection = Read-Host "Enter the number of the script to run (or type 'exit' to quit)"
    
    if ($selection -eq 'exit') {
        return
    }
    
    if ($selection -as [int] -and $selection -le $scripts.Count) {
        $scriptToRun = $scripts[$selection - 1].Url
        Write-Host "Downloading and running $($scripts[$selection - 1].Name)..."

        $scriptContent = Invoke-RestMethod -Uri $scriptToRun
        Invoke-Expression $scriptContent
    } else {
        Write-Host "Invalid selection. Please try again."
        Show-ScriptMenu
    }
}

# Show the menu
Show-ScriptMenu
