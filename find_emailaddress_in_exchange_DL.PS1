# Connect to Exchange Online using modern authentication
Connect-ExchangeOnline -UserPrincipalName ADMIN EMAIL ADDRESS GOES HERE
# Specify the email address to search for
$emailAddress = "USERSEMAIL@ADDRESS.GOES.HERE"

# Get all distribution groups
$distributionGroups = Get-DistributionGroup -ResultSize Unlimited

# Initialize an empty array to hold the groups the user is a member of
$userGroups = @()

# Loop through each distribution group and check if the user is a member
foreach ($group in $distributionGroups) {
    $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited
    foreach ($member in $members) {
        if ($member.PrimarySmtpAddress -eq $emailAddress) {
            $userGroups += $group.Name
        }
    }
}

# Display the results
if ($userGroups.Count -eq 0) {
    Write-Output "The user is not a member of any distribution group."
} else {
    Write-Output "The user is a member of the following distribution groups:"
    $userGroups | ForEach-Object { Write-Output $_ }
}
