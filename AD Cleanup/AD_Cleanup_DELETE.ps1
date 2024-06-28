##  Active Directory Cleanup Script
##  This script deletes all AD objects in the Leavers OU, based on certain criteria :
# Are currently disabled
# have not been been modified in the last 60 days
##  Avi Gondker 2024

## Suggestion for improvments
# Variable for number of days to use
# Variable for the path
# Variable for the OU
# Add a GUI?

##  Revisions
##  1/6/2024 - V1 First Draft

# Import the Active Directory module
#Import-Module ActiveDirectory

# Get the date 60 days ago
$date = (Get-Date).AddDays(-60)

# Get all AD objects that are disabled and have not been modified in the last 60 days
$adObjects = Get-ADUser -SearchBase "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk" -Filter 'Enabled -eq $false -and Modified -lt $date' -server sbs-az019105.sbs.ox.ac.uk

foreach ($object in $adObjects) {
    Delete the AD object
    Remove-ADObject -Identity $object.ObjectGUID -Confirm:$false
}