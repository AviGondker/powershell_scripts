##  Active Directory Cleanup Report Generator
##  This script exports a list of all DISABLED accounts in the Leavers OU, that have not been modified for 60 days.
##  Avi Gondker 2024

## Suggestion for improvments
# Variable for number of days to use
# Variable for the path
# Variable for the OU
# Add a GUI?

#  Revisions
#  1/6/2024 - V1   First Draft

#Import-Module ActiveDirectory

# Get the date 60 days ago
$date = (Get-Date).AddDays(-60)

# Get all AD users that are DISABLED and have not been modified in the last 60 days
Get-ADUser -SearchBase "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk" -Filter 'Enabled -eq $false' -Properties * -server sbs-az019105 | Where-Object {$_.Modified -lt $date} |Select Created, Modified, Enabled, DisplayName, EmailAddress,LastLogonDate | export-csv c:\temp\leavers_cleanup.csv -NoType