# Install and import the Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
Import-Module ExchangeOnlineManagement

# Connect to the Security & Compliance Center
$UserCredential = Get-Credential
Connect-IPPSSession -Credential $UserCredential

# Define the mailbox and search criteria
$Mailbox = "USERSNAME"
$SearchQuery = 'Subject:"AND MEETING NAME"'
$SearchName = "DeleteSubjectANDMEETINGNAME"

# Create a compliance search
New-ComplianceSearch -Name $SearchName -ExchangeLocation $Mailbox -ContentMatchQuery $SearchQuery

# Start the compliance search
Start-ComplianceSearch -Identity $SearchName

# Wait for the search to complete (you might need a longer wait time depending on the mailbox size)
Start-Sleep -Seconds 60

# Create and start a compliance search action to purge the search results
New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete

# Disconnect from Security & Compliance Center
Disconnect-ExchangeOnline -Confirm:$false
