##  Student Account Creation Script V1.0
##  This script imports a list of users details from a CSV file, and updates the Mobile attribute to their new Mobile number
##  Avi Gondker 2021

## Revisions
## 30/11 - V1.0 Initial Draft

Import-module activedirectory

## Sets a few initial variables
$date = Get-Date -Format dd-mm-yy-hhmm
$ErrorLog = @()
$SuccessLog = @()
$VerbosePreference = "Continue"
$LogFolder = "C:\Phone_Changes\Logs"

## Starts Session logging
Start-Transcript -Path "$LogFolder\Phone_ADUpdate-$date.log"

## Import the CSV File.
## This needs columns with headers Username, UserPrincipalName and Mobile
$users = Import-Csv -Path c:\phone_changes\mobile_update.csv

## Loops though the imported CSV. If the UPN can be matched, the Mobile Attribute is changed, else the error is recorded in the logs
## Attribute to search on, or change, can be adjusted, but must be included in the CSV file
## Multiple attributes can be set at the saem time, if required (will need tweak to script)
ForEach($user in $users)
{
$ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($user.UserPrincipalName)'"
     if ($ADUser)
     {
       Set-ADUser -Identity $user.username -MobilePhone $user.Mobile
       Write-Verbose "[PASS] Updated Phone number for [$($user.UserPrincipalName)]"
       $global:SuccessLog += $user.UserPrincipalName + "- Phone Number Updated sucessfully"
       
     }
     else
   {
        Write-Warning "[WARNING] UserPrincipalName for [$($user.UserPrincipalName)] does not exist"
        $global:ErrorLog += $user.UserPrincipalName + "- Cannot find user - Phone attribute NOT updated!!!"
                                     
       }

## Writes Log files to Disk
$ErrorLog | out-file -FilePath  $LogFolder\fail-$date.log -Force
$SuccessLog | out-file -FilePath  $LogFolder\success-$date.log -Force
}

Stop-Transcript
