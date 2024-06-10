<############################################################################################################

Purpose: Off-loading employees in both Active Directory and Exchange.

Chain:

Active Directory Section:
* Asks admin for a user name to disable.
* Checks for active user with that name.
* Disables user in AD.
* Resets the password of the user's AD account.
* Adds the path of the OU that the user came from to the "Description" of the account.
* Strips group memberships from user's AD account.
* Moves user's AD account to the "Disabled Users" OU.

## Suggestions for developments :
# DONE - Implement GUI - enabled users to enter username and ticket ref
# Implement a tick box for "Holding" - will then move to Holding and add additional/different notes etc?
# Exports a list of the user's group memberships (permissions) to an Excel file in a specified directory?

Version History
V1 - Script Only
V2 - 03/05/2024 - Added a GUI

############################################################################################################>

Import-Module ActiveDirectory

# Initialise Variables
$LogFolder = "C:\offboarding\logs"
$VerbosePreference = "Continue"
$ActionLog = @()

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

# Add the Exit Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(140,240)
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
$user = $(try {Get-ADUser $sam -properties distinguishedName, displayName} catch {$null})
$din = $user.displayName

If ($user -ne $Null) {

    $Account_confirmLabel.Text = "Found account for $din, $sam"
} Else {
    $Account_confirmLabel.Text =  "WARNING : Could not find account for $din, $sam"
}
}
)

#What happens when you click the Process Account button
$processButton.Add_Click(
{

#Starts Session logging 
#Uncomment for debugging
#Start-Transcript -Path "$LogFolder\session-$date.log"#Starts Session logging

$sam = $UsernameField.Text
$user = $(try {Get-ADUser $sam -properties distinguishedName, displayName} catch {$null})
$din = $user.displayName
$dn = $user.distinguishedName
$ticketRef = $TicketField.Text
$date = [datetime]::Today.ToString('dd-MM-yyyy')

# Set Account Expiry Date
Set-ADAccountExpiration -Identity $dn -DateTime $date
Write-Verbose  ("* " + $din + "'s Active Directory Account set to expire")
$global:ActionLog += $din + " Active Directory account expiry set"

# Disable the account
Disable-ADAccount $dn
Write-Verbose ($din + "'s Active Directory account is disabled.")
$global:ActionLog += $din + " Account Disabled"

# Strip the permissions from the account
Get-ADUser $dn -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object {Remove-ADGroupMember $_ -member $dn -Confirm:$false} 
Write-Verbose  ("* " + $din + "'s Active Directory group memberships (permissions) stripped from account")
$global:ActionLog += $din + " Active Directory group memberships (permissions) stripped from account"

# Add the relevant info to the leavers description on the account's properties page
Set-ADUser $dn -Description ("Leaver - $ticketRef - $date")

# Uncomment the line below to clears the Line Manager field
#Set-ADUser -Identity $dn -Clear Manager

Write-Verbose  ("* " + $din + "'s Active Directory Description updated.")
$global:ActionLog += $din + " Attributes Updated - Description"

# Move the account to the Leavers OU
Move-ADObject -Identity $dn -TargetPath "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk"
Write-Verbose  ("* " + $din + "'s Active Directory account moved to 'Leavers' OU")
$global:ActionLog += $din + " Active Directory account moved to 'Leavers' OU"

$global:ActionLog | out-file -FilePath  $LogFolder\DisableAccount-$date.log -Force

#Uncomment if you have re-enabled session logging for debugging
#Stop-Transcript

[System.Windows.Forms.MessageBox]::Show("Account for $din has been processed")
}
)

$main_form.ShowDialog()