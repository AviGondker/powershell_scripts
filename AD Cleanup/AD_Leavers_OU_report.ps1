# Import the Active Directory module
Import-Module ActiveDirectory

# Export all users in the Leavers OU
Get-ADUser -SearchBase "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk" -Filter * -Property * -server sbs-az019105 | Select Created, Modified, Enabled, DisplayName, EmailAddress,LastLogonDate | export-csv c:\temp\leavers_ou.csv -NoType