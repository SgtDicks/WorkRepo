# Set the path to the youtube-dl executable
$youtubeDlPath = "C:\path\to\youtube-dl.exe"

# Set the URL of the YouTube channel
$channelUrl = "https://www.youtube.com/c/CHANNEL_NAME"

# Set the output directory
$outputDir = "C:\YouTube Videos"

# Run youtube-dl with specified options to get the title of the latest video
$latestVideoTitle = & $youtubeDlPath --get-filename --skip-download --playlist-end 1 $channelUrl

# Create the full path of the video file to check if it exists
$videoFilePath = Join-Path -Path $outputDir -ChildPath $latestVideoTitle

if (-Not (Test-Path $videoFilePath)) {
    # Download the latest video if it doesn't exist in the output directory
    & $youtubeDlPath --format best --output "$outputDir\%(title)s.%(ext)s" --playlist-end 1 $channelUrl
} else {
    Write-Host "The latest video has already been downloaded."
}
