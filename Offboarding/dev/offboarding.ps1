<############################################################################################################

Purpose: Off-loading employees in both Active Directory and Exchange.

Chain:

Active Directory Section:
* Asks admin for a user name to disable.
* Checks for active user with that name.
* Disables user in AD.
* Resets the password of the user's AD account.
* Adds the path of the OU that the user came from to the "Description" of the account.
* Exports a list of the user's group memberships (permissions) to an Excel file in a specified directory.
* Strips group memberships from user's AD account.
* Moves user's AD account to the "Disabled Users" OU.

## Suggestions for developments :
Implement GUI - enabled users to enter username and ticket ref
Implement a tick box for "Holding" - will then move to Holding and add additional/different notes etc?


############################################################################################################>

# Import-Module ActiveDirectory
Import-module activedirectory

# Initialise Variables
$logpath = "C:\scripts\Offboarding\Logs\"
$date = [datetime]::Today.ToString('dd-MM-yyyy')

# Get the name of the account to disable from the admin
#User Principle Name
$sam = Read-Host 'Account name to disable'
$ticketRef = Read-Host 'Ticket Reference'

# Get the properties of the account and set variables
$user = Get-ADuser $sam -properties distinguishedName, displayName
$dn = $user.distinguishedName
$din = $user.displayName

## Starts Session logging
Start-Transcript -Path "$logpath"

# Disable the account
Disable-ADAccount $dn
Write-Verbose ($din + "'s Active Directory account is disabled.")

# Add the relevant info to the leavers description on the account's properties page, clean out manager etc
Set-ADUser $dn -Description ("Leaver : $ticketRef - $date")
Set-ADUser -Identity $dn -Clear Manager
Write-Verbose  ("* " + $din + "'s Active Directory Description updated.")

# Strip the permissions from the account
Get-ADUser $dn -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object {Remove-ADGroupMember $_ -member $dn -Confirm:$false} 
Write-Verbose  ("* " + $din + "'s Active Directory group memberships (permissions) stripped from account")

# Move the account to the Disabled Users OU
Move-ADObject -Identity $dn -TargetPath "OU=Leavers, OU=Disabled Accounts, OU=Decommissioned Computers, DC=homenet, DC=local"
Write-Verbose  ("* " + $din + "'s Active Directory account moved to 'Leavers' OU")
