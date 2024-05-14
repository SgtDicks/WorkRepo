# Get the current directory
$directory = Get-Location

# Prompt the user for the desired depth
$depth = Read-Host -Prompt "Enter the depth of folders to search (e.g., 3 for 3 levels deep)"

# Validate the input is a positive integer
if (-not [int]::TryParse($depth, [ref]$null) -or $depth -le 0) {
    Write-Host "Invalid input. Please enter a positive integer."
    exit
}

# Initialize an array list to store the permissions data
$permissionsData = New-Object System.Collections.ArrayList

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

        if ($CurrentDepth -gt $MaxDepth) {
            return
        }

        $folders = Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue

        foreach ($folder in $folders) {
            [PSCustomObject]@{
                FolderPath = $folder.FullName
                Depth = $CurrentDepth
            }
            Get-FoldersWithDepth -BasePath $folder.FullName -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1)
        }
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
            
            if ($acl.Access -is [System.Collections.IEnumerable]) {
                foreach ($access in $acl.Access) {
                    # Create a custom object to store the relevant information
                    $permissionsData.Add([PSCustomObject]@{
                        Folder        = $folder.FolderPath
                        Identity      = $access.IdentityReference
                        AccessControl = $access.FileSystemRights
                        Inheritance   = $access.InheritanceFlags
                        Propagation   = $access.PropagationFlags
                        Type          = $access.AccessControlType
                    }) | Out-Null
                }
            } else {
                Write-Warning "ACL Access property is not enumerable for folder: $($folder.FolderPath)"
            }
        } catch {
            Write-Warning "Failed to get ACL for folder: $($folder.FolderPath). Error: $($_.Exception.Message)"
        }
    }
}

# Call the function with the user-provided depth
Get-FolderPermissions -Path $directory -Depth $depth

# Generate collapsible HTML content
$htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Folder Permissions Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        .identity { margin: 10px 0; }
        .identity-header { cursor: pointer; padding: 10px; background: #f1f1f1; border: 1px solid #ddd; }
        .identity-content { display: none; padding: 10px; border: 1px solid #ddd; border-top: none; }
        .folder { margin: 10px 0; }
        .folder-header { font-weight: bold; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <h1>Folder Permissions Report</h1>
"@

$htmlFooter = @"
    <script>
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
        });
    </script>
</body>
</html>
"@

$htmlContent = ""

# Group the permissions by Identity
$groupedPermissions = $permissionsData | Group-Object -Property Identity

foreach ($group in $groupedPermissions) {
    $identity = $group.Name
    $foldersHtml = $group.Group | Group-Object -Property Folder | ForEach-Object {
        $folder = $_.Name
        $permissionsRows = @()
        $_.Group | ForEach-Object {
            $permissionsRows += "<tr><td>$($_.AccessControl)</td><td>$($_.Inheritance)</td><td>$($_.Propagation)</td><td>$($_.Type)</td></tr>"
        }
        $permissionsRowsString = $permissionsRows -join ""
        @"
<div class="folder">
    <div class="folder-header">$folder</div>
    <div class="folder-content">
        <table>
            <tr>
                <th>AccessControl</th>
                <th>Inheritance</th>
                <th>Propagation</th>
                <th>Type</th>
            </tr>
            $permissionsRowsString
        </table>
    </div>
</div>
"@
    }
    $foldersHtmlString = $foldersHtml -join ""
    $htmlContent += @"
<div class="identity">
    <div class="identity-header">$identity</div>
    <div class="identity-content">
        $foldersHtmlString
    </div>
</div>
"@
}

$outputHtml = Join-Path -Path $directory -ChildPath "permissions.html"

# Combine header, content, and footer
$htmlFullContent = $htmlHeader + $htmlContent + $htmlFooter

# Save the HTML content to a file
$htmlFullContent | Out-File -FilePath $outputHtml -Encoding utf8

Write-Host "Permissions report has been generated and saved to $outputHtml"
