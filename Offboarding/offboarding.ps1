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

############################################################################################################>


$date = [datetime]::Today.ToString('dd-MM-yyyy')

# Import-Module ActiveDirectory
Clear-Host
Write-Host "Offboard a user
"
# Get the name of the account to disable from the admin
$sam = Read-Host 'Account name to disable'

# Get the properties of the account and set variables
$user = Get-ADuser $sam -properties canonicalName, distinguishedName, displayName, mailNickname
$dn = $user.distinguishedName
$cn = $user.canonicalName
$din = $user.displayName
$UserAlias = $user.mailNickname

# Path building
$path1 = "C:\scripts\Offboarding\Logs\"
$path2 = "-AD-DisabledUserPermissions.csv"
$pathFinal = $path1 + $din + $path2

# Disable the account
Disable-ADAccount $sam
Write-Host ($din + "'s Active Directory account is disabled.")

# Add the relevant info to the leavers description on the account's properties page
Set-ADUser $dn -Description (Leaver' : '$ticketRef' - '$date")
Write-Host ("* " + $din + "'s Active Directory Description updated.")

# Strip the permissions from the account
Get-ADUser $User -Properties MemberOf | Select -Expand MemberOf | %{Remove-ADGroupMember $_ -member $User}
Write-Host ("* " + $din + "'s Active Directory group memberships (permissions) stripped from account")

# Move the account to the Disabled Users OU
Move-ADObject -Identity $dn -TargetPath "OU=Leavers, OU=Disabled Accounts, OU=Decommissioned Computers, DC=sbs, DC=ox, DC=ac, DC=uk"
Write-Host ("* " + $din + "'s Active Directory account moved to 'Leavers' OU")
Write-Host "Leavers Account Processes"