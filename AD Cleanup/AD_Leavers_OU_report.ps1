##  Active Directory Leavers OU Export script
##  This script exports ALL AD obejcts in the Leavers OU
##  Avi Gondker 2024

## Suggestion for improvments
# Variables for columns to use in the report
# Variable for the OU Path to use with other areas?
# Add a GUI?

##  Revisions
##  1/6/2024 - V1 First Draft

# Import the Active Directory module
#Import-Module ActiveDirectory

# Export all users in the Leavers OU
Get-ADUser -SearchBase "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk" -Filter * -Property * -server sbs-az019105 | Select Created, Modified, Enabled, DisplayName, EmailAddress,LastLogonDate | export-csv c:\temp\leavers_ou.csv -NoType