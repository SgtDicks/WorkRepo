# Prompt the user for the directory to scan (supports UNC paths)
$directory = Read-Host -Prompt "Enter the directory to scan (use UNC path for network drives, e.g., \\Server\Share)"

# Validate the directory
if (-not (Test-Path -Path $directory -PathType Container)) {
    Write-Host "Invalid directory. Please enter a valid directory path."
    exit
}

# Prompt the user for the desired depth
$depthInput = Read-Host -Prompt "Enter the depth of folders to search (e.g., 3 for 3 levels deep)"

# Initialize the depth variable
[int]$depth = 0

# Validate the depth input
if (-not [int]::TryParse($depthInput, [ref]$depth) -or $depth -le 0) {
    Write-Host "Invalid input. Please enter a positive integer."
    exit
}

# Prompt the user for the output file location
$outputPath = Read-Host -Prompt "Enter the full path for the output HTML file (e.g., C:\Reports\permissions.html)"

# Validate the output path
$outputDirectory = Split-Path -Path $outputPath -Parent
if (-not (Test-Path -Path $outputDirectory -PathType Container)) {
    Write-Host "Invalid output directory. Please enter a valid path."
    exit
}

# Initialize an array to store the permissions data
$permissionsData = New-Object System.Collections.Generic.List[PSCustomObject]

# Function to get folder permissions
function Get-FolderPermissions {
    param (
        [string]$Path,
        [int]$Depth
    )

    function Get-FoldersWithDepth {
        param (
            [string]$BasePath,
            [int]$MaxDepth,
            [int]$CurrentDepth = 0
        )

        $result = @()

        if ($CurrentDepth -gt $MaxDepth) {
            return $result
        }

        $folders = Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue

        foreach ($folder in $folders) {
            $result += [PSCustomObject]@{
                FolderPath = $folder.FullName
                Depth = $CurrentDepth
            }
            $result += Get-FoldersWithDepth -BasePath $folder.FullName -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1)
        }

        return $result
    }

    $foldersWithDepth = Get-FoldersWithDepth -BasePath $Path -MaxDepth $Depth

    # Include the root directory itself
    $foldersWithDepth += [PSCustomObject]@{
        FolderPath = (Get-Item -Path $Path).FullName
        Depth = 0
    }

    foreach ($folder in $foldersWithDepth) {
        try {
            # Get the ACL for the folder
            $acl = Get-Acl -Path $folder.FolderPath -ErrorAction Stop
            Write-Host "Processing folder: $($folder.FolderPath)"

            # Check if inheritance is enabled or broken
            $inheritanceStatus = if ($acl.AreAccessRulesProtected) { "Broken" } else { "Inherited" }

            # Create a custom object to store the relevant information
            $permissionsData.Add([PSCustomObject]@{
                Folder        = $folder.FolderPath
                Inheritance   = $inheritanceStatus
            })
        } catch {
            Write-Warning "Failed to get ACL for folder: $($folder.FolderPath). Error: $($_.Exception.Message)"
        }
    }
}

# Call the function with the user-provided depth
Get-FolderPermissions -Path $directory -Depth $depth

# Generate HTML content
$htmlHeader = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Folder Inheritance Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .folder-path { padding: 10px; }
        .inherited { color: green; }
        .broken { color: red; }
    </style>
</head>
<body>
    <h1>Folder Inheritance Report</h1>
    <div id='folderContainer'>
        <table border='1' cellspacing='0' cellpadding='5'>
            <tr>
                <th>Folder Path</th>
                <th>Inheritance Status</th>
            </tr>
"@

$htmlFooter = @"
        </table>
    </div>
    <script>
        function togglePermissions(id) {
            var content = document.getElementById(id);
            if (content.style.display === 'none') {
                content.style.display = 'block';
            } else {
                content.style.display = 'none';
            }
        }

        document.addEventListener('DOMContentLoaded', function() {
            var headers = document.querySelectorAll('.identity-header');
            headers.forEach(function(header) {
                header.addEventListener('click', function() {
                    var content = header.nextElementSibling;
                    if (content.style.display === 'block') {
                        content.style.display = 'none';
                    } else {
                        content.style.display = 'block';
                    }
                });
            });

            var brokenLinks = document.querySelectorAll('.broken');
            brokenLinks.forEach(function(link) {
                link.addEventListener('click', function() {
                    var id = link.getAttribute('data-id');
                    togglePermissions(id);
                });
            });
        });
    </script>
</body>
</html>
"@

$htmlContent = ""

foreach ($item in $permissionsData) {
    $inheritanceClass = if ($item.Inheritance -eq "Inherited") { "inherited" } else { "broken" }
    $htmlContent += "<tr>"
    $htmlContent += "<td class='folder-path'>$($item.Folder)</td>"
    $htmlContent += "<td class='$inheritanceClass'>$($item.Inheritance)</td>"
    $htmlContent += "</tr>"
}

$outputHtml = $htmlHeader + $htmlContent + $htmlFooter

# Save the HTML content to a file
$outputHtml | Out-File -FilePath $outputPath -Encoding utf8

Write-Host "Inheritance report has been generated and saved to $outputPath"
