# Import the Active Directory module
#Import-Module ActiveDirectory

# Get the date 60 days ago
$date = (Get-Date).AddDays(-60)

# Get all AD objects that are disabled and have not been modified in the last 60 days
$adObjects = Get-ADUser -SearchBase "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk" -Filter 'Enabled -eq $false -and Modified -lt $date' -server sbs-az019105.sbs.ox.ac.uk | export-csv c:\temp\info.csv -NoType
