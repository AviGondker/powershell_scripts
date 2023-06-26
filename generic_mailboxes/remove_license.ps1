##  License Removal Script V1.0
##  This script imports a list of mailboxes from a CSV file and removes ALL licenses assigned to them
##  It logs the script output to a file for review afterwards
##  Avi Gondker 2021

$VerbosePreference = "Continue"
$date = Get-Date -Format dd-mm-yy-hhmm
Connect-MsolService

## Change this patch to where your CSV file is localed, and for where you want the log file created
Start-Transcript -Path "C:\Generic_Mailboxes\logs\log-$date.log"

$users = Import-Csv -Path "C:\Generic_Mailboxes\import_files\mailboxes.csv"
foreach ($user in $users) {
try { $user = Get-MsolUser -UserPrincipalName $user.UserPrincipalName -ErrorAction Stop }
catch { continue }
 
$SKUs = @($user.Licenses)
if (!$SKUs) { Write-Verbose "NO LICENSE : No Licenses found for user $($user.UserPrincipalName), skipping..." ; continue }
 
foreach ($SKU in $SKUs) {
if (($SKU.GroupsAssigningLicense.Guid -ieq $user.ObjectId.Guid) -or (!$SKU.GroupsAssigningLicense.Guid)) {
Write-Verbose "SUCCESSSFUL: Removed license $($Sku.AccountSkuId) for user $($user.UserPrincipalName)"
Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $SKU.AccountSkuId
}
else {
Write-Verbose "GROUP: License $($Sku.AccountSkuId) for user $($user.UserPrincipalName) is assigned via Group, use the Azure AD blade to remove it!"
continue
}
}
}
Stop-Transcript