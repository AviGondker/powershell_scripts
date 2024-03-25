
# Import-Module ActiveDirectory
Import-module activedirectory

# Initialise Variables
$logpath = "C:\Offboarding\Logs\"
$date = [datetime]::Today.ToString('dd-MM-yyyy')
$ActionLog = @()
$VerbosePreference = "Continue"

# Get the name of the account to disable from the admin
#User Principle Name
$sam = Read-Host 'Account name'

# Get the properties of the account and set variables
$user = Get-ADuser $sam -properties distinguishedName, displayName
$dn = $user.distinguishedName
$din = $user.displayName

#Starts Session logging
Start-Transcript -Path "$LogFolder\session-$date.log"

# Disable the account
Set-ADUser -Identity $sam -PasswordExpired:$FALSE}
Write-Verbose ($din + "'s Active Directory account is disabled.")
$ActionLog += $user.username + " Account Password Expiry set to False"


$ActionLog | out-file -FilePath  $LogFolder\DisableAaccount-$date.log -Force
Stop-Transcript
