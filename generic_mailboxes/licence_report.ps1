## 09/11/2021 - Avi Gondker
##  This script retrieves license type and apps assigned to all users in the organisation

## When run, a login box will appear - enter credentials of a user who has access to Exchange Online (ie SBS adm accounts)
Import-Module MSOnline 
#$cred = Get-Credential 
#Connect-MsolService -Credential $cred 

Connect-MsolService

$Report = @()

$userUPN = Import-Csv -Path "C:\Generic_Mailboxes\import_files\mailboxes-SBS.csv"

foreach($User in $UserUPN)
{

    $Licenses = (Get-MsolUser -UserPrincipalName $User.UserPrincipalName).Licenses
    
    foreach($License in $Licenses)
    {
        $ServicePlans = $License.ServiceStatus

        foreach($ServicePlan in $ServicePlans)
        {
        
            $properties = @{
            Displayname = $User.Displayname
            UserPrincipalName = $User.UserPrincipalName
            Licenses = $License.AccountSkuId
            ServicePlan = $ServicePlan.ServicePlan.ServiceName
            ServicePlansStatus = $ServicePlan.ProvisioningStatus
            }
            $Report += New-Object psobject -Property $properties
        }
    }
}
# This section will create a CSV file with all the information requsted
$Report | Select Displayname, UserPrincipalName, CanonicalName, Licenses, ServicePlan, ServicePlansStatus | Export-Csv -Path "C:\Generic_Mailboxes\reports\SBS_test_report.csv" -NoType


