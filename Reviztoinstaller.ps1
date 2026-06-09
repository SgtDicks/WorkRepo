# Download and silently install Revizto
# Run PowerShell as Administrator

$DownloadUrl = "https://update.revizto.com/v5/msi64"
$DownloadDir = "C:\apps"
$MsiPath     = Join-Path $DownloadDir "Revizto_x64.msi"
$LogPath     = Join-Path $DownloadDir "ReviztoInstall.log"

# Create download folder
if (!(Test-Path $DownloadDir)) {
    New-Item -Path $DownloadDir -ItemType Directory -Force | Out-Null
}

Write-Host "Downloading Revizto installer..." -ForegroundColor Cyan

try {
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $MsiPath -UseBasicParsing
}
catch {
    Write-Error "Failed to download Revizto installer. Error: $($_.Exception.Message)"
    exit 1
}

# Confirm MSI downloaded
if (!(Test-Path $MsiPath)) {
    Write-Error "MSI was not downloaded to $MsiPath"
    exit 1
}

Write-Host "Installing Revizto silently..." -ForegroundColor Cyan

$Arguments = @(
    "/i"
    "`"$MsiPath`""
    "/q"
    "/qn"
    "/l*v"
    "`"$LogPath`""
)

$Process = Start-Process -FilePath "msiexec.exe" -ArgumentList $Arguments -Wait -PassThru

if ($Process.ExitCode -eq 0) {
    Write-Host "Revizto installed successfully." -ForegroundColor Green
    Write-Host "Install log: $LogPath"
}
else {
    Write-Error "Revizto install failed. Exit code: $($Process.ExitCode). Check log: $LogPath"
    exit $Process.ExitCode
}
