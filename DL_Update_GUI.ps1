Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#-------------------------
# Create the main form
#-------------------------
$form                  = New-Object System.Windows.Forms.Form
$form.Text            = "Distribution List Manager"
$form.Size            = New-Object System.Drawing.Size(700,600)
$form.StartPosition   = "CenterScreen"

# Flag to track if we've already connected to Exchange
$global:ConnectedToExchange = $false

function Connect-ExoIfNeeded {
    if (-not $global:ConnectedToExchange) {
        try {
            Connect-ExchangeOnline -ErrorAction Stop
            $global:ConnectedToExchange = $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to connect to Exchange Online: $($_.Exception.Message)","Error","OK","Error")
        }
    }
}

#-------------------------
# Label & TextBox: DL Email
#-------------------------
$lblDL = New-Object System.Windows.Forms.Label
$lblDL.Text = "Distribution List Email:"
$lblDL.Location = New-Object System.Drawing.Point(20,20)
$lblDL.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($lblDL)

$txtDL = New-Object System.Windows.Forms.TextBox
$txtDL.Location = New-Object System.Drawing.Point(180,20)
$txtDL.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($txtDL)

#-------------------------
# Button: Load Current Members
#-------------------------
$btnLoadMembers = New-Object System.Windows.Forms.Button
$btnLoadMembers.Text = "Load Members"
$btnLoadMembers.Location = New-Object System.Drawing.Point(500,20)
$btnLoadMembers.Size = New-Object System.Drawing.Size(100,25)
$form.Controls.Add($btnLoadMembers)

#-------------------------
# Label: Current Members
#-------------------------
$lblCurrentMembers = New-Object System.Windows.Forms.Label
$lblCurrentMembers.Text = "Current DL Members:"
$lblCurrentMembers.Location = New-Object System.Drawing.Point(20,60)
$lblCurrentMembers.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($lblCurrentMembers)

#-------------------------
# ListBox: Current Members
#-------------------------
$lstCurrentMembers = New-Object System.Windows.Forms.ListBox
$lstCurrentMembers.Location = New-Object System.Drawing.Point(20,80)
$lstCurrentMembers.Size = New-Object System.Drawing.Size(640,150)
$lstCurrentMembers.SelectionMode = "MultiExtended"
$form.Controls.Add($lstCurrentMembers)

#-------------------------
# Button: Remove Selected
#-------------------------
$btnRemoveSelected = New-Object System.Windows.Forms.Button
$btnRemoveSelected.Text = "Remove Selected"
$btnRemoveSelected.Location = New-Object System.Drawing.Point(20,240)
$btnRemoveSelected.Size = New-Object System.Drawing.Size(120,25)
$form.Controls.Add($btnRemoveSelected)

#-------------------------
# Label & TextBoxes: New Contacts
#-------------------------
$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = "Name:"
$lblName.Location = New-Object System.Drawing.Point(20,280)
$lblName.Size = New-Object System.Drawing.Size(50,20)
$form.Controls.Add($lblName)

$txtName = New-Object System.Windows.Forms.TextBox
$txtName.Location = New-Object System.Drawing.Point(70,280)
$txtName.Size = New-Object System.Drawing.Size(250,20)
$form.Controls.Add($txtName)

$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text = "Email:"
$lblEmail.Location = New-Object System.Drawing.Point(340,280)
$lblEmail.Size = New-Object System.Drawing.Size(50,20)
$form.Controls.Add($lblEmail)

$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Location = New-Object System.Drawing.Point(390,280)
$txtEmail.Size = New-Object System.Drawing.Size(270,20)
$form.Controls.Add($txtEmail)

#-------------------------
# ListBox: New Contacts
#-------------------------
$lblNewContacts = New-Object System.Windows.Forms.Label
$lblNewContacts.Text = "New Contacts to Add:"
$lblNewContacts.Location = New-Object System.Drawing.Point(20,320)
$lblNewContacts.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($lblNewContacts)

$lstNewContacts = New-Object System.Windows.Forms.ListBox
$lstNewContacts.Location = New-Object System.Drawing.Point(20,340)
$lstNewContacts.Size = New-Object System.Drawing.Size(640,120)
$form.Controls.Add($lstNewContacts)

#-------------------------
# Button: Add Contact
#-------------------------
$btnAddContact = New-Object System.Windows.Forms.Button
$btnAddContact.Text = "Add Contact"
$btnAddContact.Location = New-Object System.Drawing.Point(20,470)
$btnAddContact.Size = New-Object System.Drawing.Size(100,25)
$form.Controls.Add($btnAddContact)

#-------------------------
# Button: Process New Contacts
#-------------------------
$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = "Create & Add Contacts"
$btnProcess.Location = New-Object System.Drawing.Point(140,470)
$btnProcess.Size = New-Object System.Drawing.Size(140,25)
$form.Controls.Add($btnProcess)

#-------------------------
# FUNCTIONS / EVENT HANDLERS
#-------------------------

# Load Members button
$btnLoadMembers.Add_Click({
    $lstCurrentMembers.Items.Clear()
    if ([string]::IsNullOrWhiteSpace($txtDL.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Distribution List email first.","Error","OK","Error")
        return
    }
    Connect-ExoIfNeeded

    try {
        $members = Get-DistributionGroupMember -Identity $txtDL.Text -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to get members of $($txtDL.Text). Error: $_","Error","OK","Error")
        return
    }

    foreach ($m in $members) {
        # We can display "Name <PrimarySmtpAddress>"
        # or fallback to ExternalEmailAddress if it's a MailContact
        $emailProp = if ($m.ExternalEmailAddress) {
            $m.ExternalEmailAddress.ToString().Replace("SMTP:", "")
        } else {
            $m.PrimarySmtpAddress
        }
        $display = "$($m.Name) <$($emailProp)>"
        
        # We can store the entire object or just the email in the Tag
        $item = New-Object System.Windows.Forms.ListViewItem($display)
        # We'll store the object in Tag so we can remove it accurately
        $item.Tag = $m
        $lstCurrentMembers.Items.Add($display) | Out-Null
    }
})

# Remove Selected button
$btnRemoveSelected.Add_Click({
    if ($lstCurrentMembers.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select one or more members to remove.","Info","OK","Information")
        return
    }
    Connect-ExoIfNeeded

    # We need to remove each selected item from the DL
    # Because we used just strings in the list box, we have to parse the email.
    # Alternatively, if we store the entire object, we'd do it differently.
    foreach ($selectedItem in $lstCurrentMembers.SelectedItems) {
        # $selectedItem is just a string like "Name <email>"
        $selectedEmail = $selectedItem -replace ".*<(.*)>.*",'$1'
        try {
            Remove-DistributionGroupMember -Identity $txtDL.Text -Member $selectedEmail -Confirm:$false -ErrorAction Stop
            $lstCurrentMembers.Items.Remove($selectedItem)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to remove $selectedItem from $($txtDL.Text). Error: $_","Error","OK","Error")
        }
    }
})

# Add Contact button
$btnAddContact.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtName.Text) -or [string]::IsNullOrWhiteSpace($txtEmail.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter both Name and Email.","Error","OK","Error")
        return
    }
    $lstNewContacts.Items.Add("$($txtName.Text) <$($txtEmail.Text)>") | Out-Null
    $txtName.Clear()
    $txtEmail.Clear()
})

# Process New Contacts button
$btnProcess.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDL.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Distribution List email first.","Error","OK","Error")
        return
    }
    if ($lstNewContacts.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No new contacts to process.","Info","OK","Information")
        return
    }
    Connect-ExoIfNeeded

    foreach ($item in $lstNewContacts.Items) {
        # item is "Name <Email>"
        $regex = [regex]'^(?<name>.+)\s+<(?<email>.+)>$'
        $matches = $regex.Match($item)
        $name = $matches.Groups["name"].Value
        $email = $matches.Groups["email"].Value

        try {
            # Create MailContact
            New-MailContact -Name $name -ExternalEmailAddress $email -ErrorAction Stop
        } catch {
            # If contact already exists or creation fails, handle it
            Write-Host "New-MailContact failed for $name ($email): $_"
        }

        # Now add to the DL
        try {
            Add-DistributionGroupMember -Identity $txtDL.Text -Member $email -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to add $name <$email> to $($txtDL.Text). Error: $_","Error","OK","Error")
        }
    }

    [System.Windows.Forms.MessageBox]::Show("New contacts have been created (if needed) and added to the DL.","Success","OK","Information")
    $lstNewContacts.Items.Clear()
})

#-------------------------
# Show the form
#-------------------------
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
