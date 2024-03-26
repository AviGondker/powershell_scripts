Import-Module ActiveDirectory

# Initialise Variables
$LogFolder = "C:\temp\Logs\"
$date = [datetime]::Today.ToString('dd-MM-yyyy')
$VerbosePreference = "Continue"

#Creates the interface
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Offboarding Script'
$main_form.Width = 425
$main_form.Height = 400
$main_form.AutoSize = $false
$main_form.StartPosition = 'CenterScreen'
$main_form.Topmost = $true

#Add the Ticket Reference label text
$TicketText = New-Object System.Windows.Forms.Label
$TicketText.Text = "Please enter the ServiceDesk Ticket Ref below"
$TicketText.Location  = New-Object System.Drawing.Point(10,30)
$TicketText.AutoSize = $true
$main_form.Controls.Add($TicketText)

#Add the Ticket Reference text Field
$TicketField = New-Object System.Windows.Forms.TextBox
$TicketField.Location = New-Object System.Drawing.Point(10,50)
$TicketField.Size = New-Object System.Drawing.Size(180,55)
$TicketField.Multiline = $false
$main_form.Controls.Add($TicketField)

# Add the FindUser Button
$finduserButton = New-Object System.Windows.Forms.Button
$finduserButton.Location = New-Object System.Drawing.Point(250, 80)
$finduserButton.Size = New-Object System.Drawing.Size(100,50)
$finduserButton.Text = 'FindUser'
$main_form.Controls.Add($finduserButton)

#Add the Username label text
$UsernameText = New-Object System.Windows.Forms.Label 
$UsernameText.Text = "Please enter the Username below, eg ' auser '"
$UsernameText.Location  = New-Object System.Drawing.Point(10,90)
$UsernameText.AutoSize = $true
$main_form.Controls.Add($UsernameText)

#Add the Username Text field
$UsernameField = New-Object System.Windows.Forms.TextBox
$UsernameField.Location = New-Object System.Drawing.Point(10,110)
$UsernameField.Size = New-Object System.Drawing.Size(180,55)
$UsernameField.Multiline = $false
$main_form.Controls.Add($UsernameField)

#Add the Account Confirmation
$Account_confirmLabel = New-Object System.Windows.Forms.Label
$Account_confirmLabel.Text = "Account you have selected : "
$Account_confirmLabel.Location  = New-Object System.Drawing.Point(10,150)
$Account_confirmLabel.AutoSize = $true
$main_form.Controls.Add($Account_confirmLabel)

# Process the Account Button
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(10,180)
$processButton.Size = New-Object System.Drawing.Size(380,40)
$processButton.Text = "Process Account"
$main_form.Controls.Add($processButton)

## Add the View Logs Button
$viewLogsButton = New-Object System.Windows.Forms.Button
$viewLogsButton.Location = New-Object System.Drawing.Point(10,240)
$viewLogsButton.Size = New-Object System.Drawing.Size(125,50)
$viewLogsButton.Text = 'View Logs...'
$main_form.Controls.Add($viewLogsButton)

# Add the Exit Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(240,240)
$exitButton.Size = New-Object System.Drawing.Size(150,50)
$exitButton.Text = 'EXIT'
$exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$main_form.CancelButton = $exitButton
$main_form.Controls.Add($exitButton)

## Add the Copyright Notices
$CopyrightLabel = New-Object System.Windows.Forms.Label
$CopyrightLabel.Text = "Copyright 2023, Avi Gondker, All Rights Reserved"
$CopyrightLabel.AutoSize = $true
$CopyrightLabel.Location  = New-Object System.Drawing.Point(10,320)
$main_form.Controls.Add($CopyrightLabel)

# What happens when you click the Find User button
$finduserButton.Add_Click(
{
$sam = $UsernameField.Text
$user = $(try {Get-ADUser $sam -properties distinguishedName, displayName -server dc01.homenet.local} catch {$null})
$dn = $user.distinguishedName
$din = $user.displayName

If ($user -ne $Null) {

    $Account_confirmLabel.Text = "Found account for $din"
} Else {
    $Account_confirmLabel.Text =  "WARNING : Could not find account for $sam"
}
}
)

###TESTING AREA###
## 
#Account detection
$Account_testLabel = New-Object System.Windows.Forms.Label
$Account_testLabel.Text = "Account you have selected : "
$Account_testLabel.Location  = New-Object System.Drawing.Point(10,300)
$Account_testLabel.AutoSize = $true
$main_form.Controls.Add($Account_testLabel)

#What happens when you click the Process Account button
$processButton.Add_Click(
{
$sam = $UsernameField.Text
$user = $(try {Get-ADUser $sam -properties distinguishedName, displayName -server dc01.homenet.local} catch {$null})
$dn = $user.distinguishedName
$din = $user.displayName
$ticketRef = $TicketField.Text
#Starts Session logging
Start-Transcript -Path "$LogFolder\session-$date.log"

# Disable the account
Disable-ADAccount $dn
Write-Verbose ($din + "'s Active Directory account is disabled.")
#$ActionLog += $ad_user.username + " Account Disabled"

# Add the relevant info to the leavers description on the account's properties page, clean out manager etc
Set-ADUser $dn -Description ("Leaver : $ticketRef - $date")
Set-ADUser -Identity $dn -Clear Manager
Write-Verbose  ("* " + $din + "'s Active Directory Description updated.")
#$ActionLog += $ad_user.username + " Attributes Updated - Description"

# Remove the LoginScript attributes
Set-ADUser -Identity $dn -Clear ScriptPath
Write-Verbose  ("* " + $din + "'s Active Directory Loginscript removed.")
#$ActionLog += $ad_user.username + " Attributes Updated - Loginscript"

# Strip the permissions from the account
Get-ADUser $dn -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object {Remove-ADGroupMember $_ -member $dn -Confirm:$false} 
Write-Verbose  ("* " + $din + "'s Active Directory group memberships (permissions) stripped from account")
#$ActionLog += $user.username + " Active Directory group memberships (permissions) stripped from account"

# Set Account Expiry Date
Set-ADAccountExpiration -Identity $dn -DateTime $date
Write-Verbose  ("* " + $din + "'s Active Directory Account set to expire")
#$ActionLog += $user.username + " Active Directory accout set to expire"

# Move the account to the Disabled Users OU
Move-ADObject -Identity $dn -TargetPath "OU=Leavers, OU=Disabled Accounts, OU=Decommissioned Computers, DC=homenet, DC=local"
Write-Verbose  ("* " + $din + "'s Active Directory account moved to 'Leavers' OU")
#$SuccessLog += $user.username + "Active Directory account moved to 'Leavers' OU"

$ActionLog | out-file -FilePath  $LogFolder\DisableAaccount-$date.log -Force
Stop-Transcript
[System.Windows.Forms.MessageBox]::Show("Account for $din has been processed. Please review Logs if errors appear")
}
)

$main_form.ShowDialog()