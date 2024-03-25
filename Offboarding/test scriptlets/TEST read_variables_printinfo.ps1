Import-module activedirectory

$logpath = "C:\scripts\Offboarding\Logs\log.log"
$date = [datetime]::Today.ToString('dd-MM-yyyy')

Start-Transcript -Path $logpath

$sam = Read-Host 'Account name to disable'
$ticketRef = Read-Host 'Ticket Reference'

$user = Get-ADuser $sam -properties distinguishedName, displayName
$dn = $user.distinguishedName
$din = $user.displayName

Write-Host ("Ticket Ref has been added")
Write-Host ($din + "'s Active Directory account is disabled.")

echo $sam
echo $ticketref
echo $dn
echo $din