# # Experience Cloud ID Active Directory Tool V1.5
# This script adds the Salesforce Experience Cloud ID to the relevant Active Directory Accounts 
# Avi Gondker 2021


# Revisions
# 09/09/21 - V1 - First Draft - stand-alone scrpit to create MBA accounts from CSV
# 06/06/22 - V1.2 - Amended to use SBS folders, rather than local test, and uploaded to RDWEB
# 07/06/22 - V1.3 - Amended logs name, re-designed interface, added copyright notices
# 15/06/22 - V1.4 - Added requirement to add account to the "SBS-UG-ExperienceCloud" AD access group
# 11/07/22 - V1.5 - Amended AD Access group to "SBS-UG-SalesForceNewExperienceCloud"


Import-Module ActiveDirectory

#Sets up some initial parameters
$date = Get-Date -Format dd-mm-yy
$f=0
$s=0
$count=0

#Sets up the parameters for the log files
$ErrorLog = @()
$SuccessLog = @()
$VerbosePreference = "Continue"
$LogFolder = "C:\EC_Upload\logs"

#Creates the interface
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text =’Student EC ID Import Script'
$main_form.Width = 425
$main_form.Height = 400
$main_form.AutoSize = $false
$main_form.StartPosition = 'CenterScreen'
#$main_form.Topmost = $true


#Add Import File label
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select the CSV Import File"
$Label.Location  = New-Object System.Drawing.Point(10,30)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

#Add the Browse Button
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Point(10,50)
$BrowseButton.Size = New-Object System.Drawing.Size(380,40)
$BrowseButton.Text = "Browse...."
$main_form.Controls.Add($BrowseButton)

#Add the Import File Confirmation
$ImportFileLabel = New-Object System.Windows.Forms.Label
$ImportFileLabel.Text = "Selected Import File : $importfile"
$ImportFileLabel.Location  = New-Object System.Drawing.Point(10,100)
$ImportFileLabel.AutoSize = $true
$main_form.Controls.Add($ImportFileLabel)

#Add the Add Experience Cloud IDs Button
$createButton = New-Object System.Windows.Forms.Button
$createButton.Location = New-Object System.Drawing.Point(10,130)
$createButton.Size = New-Object System.Drawing.Size(380,40)
$createButton.Text = "Upload Experience Cloud ID's"
$main_form.Controls.Add($createButton)

# Add the Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Name = 'ProgressBar'
$progressBar.Style = "Continuous"
$progressBar.Location = New-Object System.Drawing.Size (10,180)
$progressBar.Size = New-Object System.Drawing.Size (380,30)
$progressBar.Value=0
$progressBar.Step=1
$progressBar.Visible=$false
$main_form.Controls.Add($progressBar)

## Add the View Logs Button
$viewLogsButton = New-Object System.Windows.Forms.Button
$viewLogsButton.Location = New-Object System.Drawing.Point(10,220)
$viewLogsButton.Size = New-Object System.Drawing.Size(125,50)
$viewLogsButton.Text = 'View Logs...'
$main_form.Controls.Add($viewLogsButton)

# Add the Exit Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(240,220)
$exitButton.Size = New-Object System.Drawing.Size(150,50)
$exitButton.Text = 'EXIT'
$exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$main_form.CancelButton = $exitButton
$main_form.Controls.Add($exitButton)

## Add the Copyright Notices
$CopyrightLabel = New-Object System.Windows.Forms.Label
$CopyrightLabel.Text = "Copyright © 2022, Avi Gondker, All Rights Reserved"
$CopyrightLabel.AutoSize = $true
$CopyrightLabel.Location  = New-Object System.Drawing.Point(10,320)
$main_form.Controls.Add($CopyrightLabel)

#What happens when you click the Browse button
$BrowseButton.Add_Click(
{
$importfile = Get-FileName "\\sbs.ox.ac.uk\groupshares$\Staff_IT\SDP Project\Students_2022\MBA 2022"
$global:users = Import-csv -Path $importfile
$ImportFileLabel.Text = "Selected Import File : $importfile"
}
)

## What Happens when you click the "View Logs" button
$viewLogsButton.Add_Click(
{
$filter = "Log Files (*.log)| *.log"
$viewfile = Get-FileName "$LogFolder" "$filter"
Invoke-Item $viewfile
}
)

# What happens when you click the Add button
$CreateButton.Add_Click(
{
#Starts Session logging
Start-Transcript -Path "$LogFolder\session-$date.log"

#Displays the Progress Bar
$progressBar.Visible=$true

#Checks for and creates logging directory if not already present
#if (!(test-path $LogFolder\$classyear)) 
#{
#   Write-Verbose "Folder [$($LogFolder)] does not exist, creating"
#   New-Item -Path $LogFolder -Name $classyear -ItemType "directory" -Force 
#   }
#Sets the Progress bar 100% value to total number of accounts in the import file

ForEach($user in $users)
{
$progressBar.Maximum=$users.Count      
#This section creates the accounts. If the AD Object does NOT already exist, create it, and add entry to the log file
try {
       Set-ADUser -Identity $user.Username -Add @{extensionAttribute5 = $user.EC_ID}
       Add-ADGroupMember -Identity "SBS-UG-SalesForceNewExperienceCloud" -Members $user.Username
       Write-Verbose "[PASS] Added EC ID code for [$($user.Username)] and added to relevant Salesforce Access group"
       $global:SuccessLog += $user.Username + " PASS - Sucesfully Added EC ID code"
       $global:s++
       $count++
       $progressBar.PerformStep()                                     
          }
      
catch {
        Write-Warning "[WARNING] SAMAccountName for [$($user.Username)] does not exist"
        $global:ErrorLog += $user.Username + " FAILED - AD object does not exist"
        $global:f++
        $count++
        $progressBar.PerformStep()                                
       }
[System.Windows.Forms.Application]::DoEvents()

#Writes the logs to disk
$ErrorLog | out-file -FilePath  $LogFolder\Import_fail-$date.log -Force
$SuccessLog | out-file -FilePath  $LogFolder\Import_success-$date.log -Force
}
Stop-Transcript
[System.Windows.Forms.MessageBox]::Show("EC IDs added!`n`nID Sucessfully Added : $s`nIDs Failed : $f  `n`nPlease review event logs for more info")
}
)

Function Get-FileName($initialDirectory, $filter)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.Title = "Please select your Import batch .csv import file"
    $OpenFileDialog.filter = $filter
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$main_form.ShowDialog()