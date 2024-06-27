Import-Module ActiveDirectory

# Get the date 60 days ago
$date = (Get-Date).AddDays(-60)

# Get all AD users that are DISABLED and have not been modified in the last 60 days
Get-ADUser -SearchBase "OU=Leavers,OU=Disabled Accounts,OU=Decommissioned Computers,DC=sbs,DC=ox,DC=ac,DC=uk" -Filter 'Enabled -eq $false' -Properties * -server sbs-az019105 | Where-Object {$_.Modified -lt $date} |Select Created, Modified, Enabled, DisplayName, EmailAddress,LastLogonDate | export-csv c:\temp\leavers_cleanup.csv -NoType