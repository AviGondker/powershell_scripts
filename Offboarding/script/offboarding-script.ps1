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
# Exports a list of the user's group memberships (permissions) to an Excel file in a specified directory.

Version History
V1 - Script Only

############################################################################################################>

# Import-Module ActiveDirectory
Import-module activedirectory

# Initialise Variables
$LogFolder = "C:\scripts\Offboarding\Logs\"
$date = [datetime]::Today.ToString('dd-MM-yyyy')
$ActionLog = @()
$VerbosePreference = "Continue"

# Get the name of the account to disable from the admin
#User Principle Name
$sam = Read-Host 'Account name to disable'
$ticketRef = Read-Host 'Ticket Reference'

# Get the properties of the account and set variables
$user = Get-ADuser $sam -properties distinguishedName, displayName
$dn = $user.distinguishedName
$din = $user.displayName

#Starts Session logging
Start-Transcript -Path "$LogFolder\session-$date.log"

# Disable the account
Disable-ADAccount $dn
Write-Verbose ($din + "'s Active Directory account is disabled.")
$ActionLog += $user.username + " Account Disabled"

# Add the relevant info to the leavers description on the account's properties page, clean out manager etc
Set-ADUser $dn -Description ("Leaver : $ticketRef - $date")
#Set-ADUser -Identity $dn -Clear Manager
Write-Verbose  ("* " + $din + "'s Active Directory Description updated.")
$ActionLog += $user.username + " Attributes Updated"

# Strip the permissions from the account
Get-ADUser $dn -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object {Remove-ADGroupMember $_ -member $dn -Confirm:$false} 
Write-Verbose  ("* " + $din + "'s Active Directory group memberships (permissions) stripped from account")
$ActionLog += $user.username + " Active Directory group memberships (permissions) stripped from account"

# Move the account to the Disabled Users OU
Move-ADObject -Identity $dn -TargetPath "OU=Leavers, OU=Disabled Accounts, OU=Decommissioned Computers, DC=homenet, DC=local"
Write-Verbose  ("* " + $din + "'s Active Directory account moved to 'Leavers' OU")
SuccessLog += $user.username + "Active Directory account moved to 'Leavers' OU"

$ActionLog | out-file -FilePath  $LogFolder\DisableAaccount-$date.log -Force
Stop-Transcript
