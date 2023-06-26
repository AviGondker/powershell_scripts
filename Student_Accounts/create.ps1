##  Student Account Creation Script V10.4
##  This script imports a list of students from a CSV file, creates the AD accounts in the specified OU, adds them to the Students_ALL group, and enables MFA
##  Avi Gondker 2021

##  Revisions
##  21/5 - V1   First Draft - stand-alone scrpit to create MBA accounts from CSV
##  23/5 - V2   Added code comments and more variables into parameter section
##  25/5 - V2.2 Added further comments, added file browser dialogue
##  26/5 - V3   Set "Change Password at next logon"
##              Removed "Never Expires"
##              Passthough variables from the Batch File to enable one-click creations
##  01/6 - V4   Automatically adds account to "Students_All" Groups
##  15/6 - V5   Added Log files for sucessful and failed imports
##  19/7 - V6   Merges all the creation sections ito one section
##  06/8 - V7   Started adding GUI
##  11/8 - V7.1 Improved GUI layout
##  12/8 - V7.2 Added Popup Window at end of script, better error handling, more cohorts 
##  13/8 - V8   SBS logo, Jan/Sept Stream handling
##  14/8 - V8.1 Cleaned up code, renamed variables, added more cohorts, fixed Jan/Sept start errors
##  17/8 - V8.2 Added Clear Form button, and changed Jan/Sept intake to Radio buttons so both cant be selected at same time
##  19/8 - v8.3 Edited Clear button behaviour, cleaned up code
##  26/8 - v8.4 Added code to handle MMPM Cohort Number in OU - eg MMPM13_Cohort13
##  03/9 - v9.0 Added Progress Bar
##  06/9 - v9.1 Added Stream 1 and 2 for DipAI courses
##  24/9 - V9.2 Added Dphils
##  4/10 - V9.3 Added MCS 1+1 Cohort, and merged Intake/Stream radio boxes
##  5/10 - V9.4 Added Undergraduates
##  6/10 - V9.5 Adjusted for correct OU format for DIPSI for Jan and March intakes
##  8/11 - V9.6 Adjusted to automatically add to the MFA enablement group "SBS-UG-Require MFA for Users (on-prem)"
##  9/11 - V9.7 Adjusted correct OU format for DIP SI stream 1 and 2 again, added info for Stream 1 /2 modifiers
##  23/11 - V9.8 Adjusted DIPSI back to Jan and March OUs
##  04/01 - v10 Added button to be able to view Log files of the created cohort 
##              new code for file filters in "Browse" to show just *.CSV, and "View Logs" to show just *.Log
##              Error messages and popup message wording adjusted
##              Added copyright notice :)
## 2022
## 26/01 - V10.1 Hide Progress bar until Create button clicked, to make interface cleaner
## 15/02 - V10.2 Adjusted for Undergraduate OUs, modified behaviour of Clear Form button to remove list box selections
## 15/12 - v10.3 Remove 2021 and 2022 from Year drop-down lists
##
## 2023
## 23/06 - V10.4 Removed 2021 and 2022 sections from MMPM cohort overrides
##               Fixed the MMPM Cohort 15 not filling in the OU correctly

Import-module activedirectory

## Sets up some initial parameters
$date = Get-Date -Format dd-mm-yy-hhmm
$classyear = ""
$year = ""
$course = ""
$cohort = ""
$visiting = ""
$oupath = ""
$f=0
$s=0
$count=0

## Sets up the parameters for the log files
$ErrorLog = @()
$SuccessLog = @()
$VerbosePreference = "Continue"
$LogFolder = "C:\Student_Accounts\logs"

## Creates the interface
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text =’Student Accounts Creation Script'
$main_form.Width = 950
$main_form.Height = 600
$main_form.AutoSize = $false
$main_form.StartPosition = 'CenterScreen'
## $main_form.Topmost = $true

## Adds the Business School Logo
$img = [System.Drawing.Image]::Fromfile('logo.png')
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Size(453,360)
$pictureBox.Width = 353
$pictureBox.Height = 173
$pictureBox.Image = $img
$main_form.controls.add($pictureBox)

## Add Instructions Part 1 to the page
$Step1 = New-Object System.Windows.Forms.Label
$Step1.Text = "Step 1"
$Step1.Location  = New-Object System.Drawing.Point(10,10)
$Step1.AutoSize = $true
$Step1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12,[System.Drawing.FontStyle]::Bold)
$main_form.Controls.Add($Step1)

$Guide1 = New-Object System.Windows.Forms.Label
$Guide1.Text = "Select the Course and Intake Year from the dropdown lists below, then click the SET button. Click the CLEAR FORM button to start again"
$Guide1.Location  = New-Object System.Drawing.Point(10,35)
$Guide1.AutoSize = $true
$main_form.Controls.Add($Guide1)

## Add a listbox for the year
$YearlistBox = New-Object System.Windows.Forms.ListBox
$YearlistBox.Location = New-Object System.Drawing.Point(300,60)
$YearlistBox.Size = New-Object System.Drawing.Size(200,20)
$YearlistBox.Height = 80

[void] $YearlistBox.Items.Add('2023')
[void] $YearlistBox.Items.Add('2024')
[void] $YearlistBox.Items.Add('2025')

$main_form.Controls.Add($yearlistBox)

$YearlistBox.add_SelectedIndexChanged(
     { 
          $Yearselected = $YearlistBox.SelectedItems
          $CourseYearLabel.Text = "Course Year : $Yearselected"
          $global:year="$Yearselected"
          }
)

## Add the Course Year writing
$CourseYearLabel = New-Object System.Windows.Forms.Label
$CourseYearLabel.Text = "Course Year : $CourseYear"
$CourseYearLabel.Location  = New-Object System.Drawing.Point(300,130)
$CourseYearLabel.AutoSize = $true
$main_form.Controls.Add($CourseYearLabel)

## Add a listbox for the Course
$CourselistBox = New-Object System.Windows.Forms.ListBox
$CourselistBox.Location = New-Object System.Drawing.Point(10,60)
$CourselistBox.Size = New-Object System.Drawing.Size(200,20)
$CourselistBox.Height = 80

[void] $CourselistBox.Items.Add('DAI')
[void] $CourselistBox.Items.Add('DFS')
[void] $CourselistBox.Items.Add('DGB')
[void] $CourselistBox.Items.Add('DOL')
[void] $CourselistBox.Items.Add('DSI')
[void] $CourselistBox.Items.Add('DPhil')
[void] $CourselistBox.Items.Add('EMBA')
[void] $CourselistBox.Items.Add('MBA')
[void] $CourselistBox.Items.Add('MFE')
[void] $CourselistBox.Items.Add('MCS1+1')
[void] $CourselistBox.Items.Add('MGHL')
[void] $CourselistBox.Items.Add('MLF')
[void] $CourselistBox.Items.Add('MPM')
[void] $CourselistBox.Items.Add('Undergraduates')


$main_form.Controls.Add($CourselistBox)

$CourselistBox.add_SelectedIndexChanged(
     { 
          $CourseSelected = $CourselistBox.SelectedItems
          $CourseLabel.Text = "Course : $CourseSelected"
          $global:course="$CourseSelected"                  
     }
)

## Add the Course Selection writing
$CourseLabel = New-Object System.Windows.Forms.Label
$CourseLabel.Text = "Course : $Course"
$CourseLabel.Location  = New-Object System.Drawing.Point(10,130)
$CourseLabel.AutoSize = $true
$main_form.Controls.Add($CourseLabel)

## Jan Intake / Stream 1 Radio Button 
$janradiobox = new-object System.Windows.Forms.RadioButton
$janradiobox.Location = new-object System.Drawing.Size(550,60)
$janradiobox.Size = new-object System.Drawing.Size(190,20)
$janradiobox.Text = "Stream 1 (EMBA, DipAI, DipSI)"
$janradiobox.Checked = $false
$main_form.Controls.Add($janradiobox) 

## Sept Intake / Stream 2 Radio Button 
$septradiobox = new-object System.Windows.Forms.RadioButton
$septradiobox.Location = new-object System.Drawing.Size(550,85)
$septradiobox.Size = new-object System.Drawing.Size(190,20)
$septradiobox.Text = "Stream 2 (EMBA, DipAI. DipSI)"
$septradiobox.Checked = $false
$main_form.Controls.Add($septradiobox)

## Add the Visiting (Undergrads) Radio Button 
$visitingradiobox = new-object System.Windows.Forms.RadioButton
$visitingradiobox.Location = new-object System.Drawing.Size(550,110)
$visitingradiobox.Size = new-object System.Drawing.Size(190,20)
$visitingradiobox.Text = "Visiting (Undergrads only)"
$visitingradiobox.Checked = $false
$main_form.Controls.Add($visitingradiobox)

## Add the SET Button
$setButton = New-Object System.Windows.Forms.Button
$setButton.Location = New-Object System.Drawing.Point(10,160)
$setButton.Size = New-Object System.Drawing.Size(75,25)
$setButton.Text = "SET"
$main_form.Controls.Add($setButton)

## Add the CLEAR Button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(100,160)
$clearButton.Size = New-Object System.Drawing.Size(85,25)
$clearButton.Text = 'Clear Form'
$main_form.Controls.Add($clearButton)

## Add the OU Confirmation line
$OUpathLabel = New-Object System.Windows.Forms.Label
$OUpathLabel.Text = "Accounts will be created in the following location:$oupath"
$OUpathLabel.Location  = New-Object System.Drawing.Point(10,200)
$OUpathLabel.AutoSize = $true
$main_form.Controls.Add($OUpathLabel)

## Add Instructions Part 2 to the page
$Step2 = New-Object System.Windows.Forms.Label
$Step2.Text = "Step 2"
$Step2.Location  = New-Object System.Drawing.Point(10,235)
$Step2.AutoSize = $true
$Step2.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12,[System.Drawing.FontStyle]::Bold)
$main_form.Controls.Add($Step2)

$Guide2 = New-Object System.Windows.Forms.Label
$Guide2.Text = "Select your import file, then click the CREATE button.The process will be complete when you see a popup window"
$Guide2.Location  = New-Object System.Drawing.Point(10,265)
$Guide2.AutoSize = $true
$main_form.Controls.Add($Guide2)

## Add import File 
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select AD Import Batch File ---> "
$Label.Location  = New-Object System.Drawing.Point(10,300)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

## Add the Browse Button
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(220,295)
$BrowseButton.Size = New-Object System.Drawing.Size(120,30)
$BrowseButton.Text = "Browse...."
$main_form.Controls.Add($BrowseButton)

## Add the Import File Confirmation
$ImportFileLabel = New-Object System.Windows.Forms.Label
$ImportFileLabel.Text = "Selected Import File : $importfile"
$ImportFileLabel.Location  = New-Object System.Drawing.Point(10,330)
$ImportFileLabel.AutoSize = $true
$main_form.Controls.Add($ImportFileLabel)

## Add the CREATE Button
$createButton = New-Object System.Windows.Forms.Button
$createButton.Location = New-Object System.Drawing.Point(10,360)
$createButton.Size = New-Object System.Drawing.Size(380,50)
$createButton.Text = 'Create...!'
$main_form.Controls.Add($createButton)

##  Add the Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Name = 'ProgressBar'
$progressBar.Style = "Continuous"
$progressBar.Location = New-Object System.Drawing.Size (10,420)
$progressBar.Size = New-Object System.Drawing.Size (380,30)
$progressBar.Value=0
$progressBar.Step=1
$progressBar.Visible=$false
$main_form.Controls.Add($progressBar)

##  Add the CLOSE Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(265,460)
$exitButton.Size = New-Object System.Drawing.Size(125,50)
$exitButton.Text = 'EXIT'
$exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$main_form.CancelButton = $exitButton
$main_form.Controls.Add($exitButton)

## Add the View Logs Button
$viewLogsButton = New-Object System.Windows.Forms.Button
$viewLogsButton.Location = New-Object System.Drawing.Point(10,460)
$viewLogsButton.Size = New-Object System.Drawing.Size(125,50)
$viewLogsButton.Text = 'View Logs...'
$main_form.Controls.Add($viewLogsButton)

## Add the Copyright Notices
$CopyrightLabel = New-Object System.Windows.Forms.Label
$CopyrightLabel.Text = "Copyright © 2021, Avi Gondker, All Rights Reserved"
$CopyrightLabel.AutoSize = $true
$CopyrightLabel.Location  = New-Object System.Drawing.Point(10,520)
$main_form.Controls.Add($CopyrightLabel)


##  What happens when you click the Set Button
## Takes the Course and Year selected to build the OU string, and prints an update on the screen of the full OU
$setButton.Add_Click(
{

##  DipSI Jan Intake/Stream 1 override
if ($global:course -eq "DSI" -and $janradiobox.Checked -eq $true)
{
$global:stream="_Jan"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year"
$global:classyear1 = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear1,$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

##  DipSI Sept Intake/Stream 2 override
elseif ($global:course -eq "DSI" -and $septradiobox.Checked -eq $true)
{
$global:stream="_March"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year"
$global:classyear1 = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear1,$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

##  EMBA Jan Intake/Stream 1 override
elseif ($janradiobox.Checked -eq $true)
{
$global:stream="_JanStart"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

##  EMBA Sept Intake/Stream 2 override
elseif ($septradiobox.Checked -eq $true)
{
$global:stream="_SeptStart"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

##   MMPM Override to deal with the Cohort No.s for different years etc
elseif  ($global:course -eq "MPM" -and $global:year -eq "2023")
{
$global:stream="_Cohort15"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "MPM" -and $global:year -eq "2024")
{
$global:stream="_Cohort16"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "MPM" -and $global:year -eq "2025")
{
$global:stream="_Cohort17"
$global:cohort = "OU=Students_$course"
$global:classyear = "OU=Students_$global:course$global:year$global:stream"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

## Dphil Overide
elseif  ($global:course -eq "Dphil")
{
$global:cohort = "OU=Students_$course"
$global:stream="_Started_Sept$global:year"
$global:classyear = "OU=Dphil$global:stream"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

## MCS 1+1 override
elseif  ($global:course -eq "MCS1+1")
{
$global:stream="_$global:year"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

## Undergrads Overrides (including visiting Undergrads)
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2021" -and $global:visitingradiobox.Checked -eq $true)
{
$global:visiting="Visiting_Undergraduates"
$global:stream="_2021-2024"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:visiting$global:stream,$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2021")
{
$global:stream="_2021-2024"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2022" -and $global:visitingradiobox.Checked -eq $true)
{
$global:visiting="Visiting_Undergraduates"
$global:stream="_2022-2025"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:visiting$global:stream,$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2022")
{
$global:stream="_2022-2025"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}

elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2023" -and $global:visitingradiobox.Checked -eq $true)
{
$global:visiting="Visiting_Undergraduates"
$global:stream="_2023-2026"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:visiting$global:stream,$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2023")
{
$global:stream="_2023-2026"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2024" -and $global:visitingradiobox.Checked -eq $true)
{
$global:visiting="Visiting_Undergraduates"
$global:stream="_2024-2027"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:visiting$global:stream,$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2024")
{
$global:stream="_2024-2027"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2025" -and $global:visitingradiobox.Checked -eq $true)
{
$global:visiting="Visiting_Undergraduates"
$global:stream="_2025-2028"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:visiting$global:stream,$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
elseif  ($global:course -eq "Undergraduates" -and $global:year -eq "2025")
{
$global:stream="_2025-2028"
$global:classyear = "OU=Students_$global:course"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear$global:stream,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
## Standard OU string for rest of cohorts
else
{
$global:classyear = "OU=Students_$global:course$global:year"
$global:cohort = "OU=Students_$course"
$global:oupath = "$global:classyear,$global:cohort,OU=Students,OU=Users,OU=SBS,DC=sbs,DC=ox,DC=ac,DC=uk"
$OUpathLabel.Text = "Accounts will be created in the following location:`n$oupath"
}
}
)
##  What Happens when you click the CLEAR button - Clears all input from the form
$clearButton.Add_Click(
{
$global:stream=""
$global:cohort = ""
$global:classyear = ""
$global:oupath = ""
$Courselistbox.ClearSelected()
$Yearlistbox.ClearSelected()
$janradiobox.Checked = $false
$septradiobox.Checked = $false
$visitingradiobox.Checked = $false
$OUpathLabel.Text = "Form has been cleared. Please select Course and Intake Year again"
}
)

##  What happens when you click the Browse Button
##  Runs the GetFilename function to open a File Browser dialogue box to select the import file
$BrowseButton.Add_Click(
{
$filter = "CSV (*.csv)| *.csv"
$importfile = Get-FileName "\\sbs.ox.ac.uk\groupshares$\Staff_IT\SDP Project" "$filter"
$global:users = Import-csv -Path $importfile
$ImportFileLabel.Text = "Selected Import File : $importfile"
}
)


## What Happens when you click the "View Logs" button
$viewLogsButton.Add_Click(
{
$filter = "Log Files (*.log)| *.log"
$viewfile = Get-FileName "$LogFolder\$classyear" "$filter"
Invoke-Item $viewfile
}
)

##  What happens when you click the Create button
##  Starts the Session Logging
##  Creates the AD accounts, based on the parameters selected
##  Runs and updates the Progress bar
##  Creates log files and also prints the final status onscreen
$CreateButton.Add_Click(
{
## Checks for and creates logging directory
if (!(test-path $LogFolder\$classyear)) 
{
    Write-Verbose "Folder [$($LogFolder)] does not exist, creating"
   New-Item -Path $LogFolder -Name $classyear -ItemType "directory" -Force 
   }

## Starts Session logging
Start-Transcript -Path "$LogFolder\$classyear\session-$date.log"

## Un-hides the Progress bar
$progressBar.Visible=$true
 
ForEach($user in $users)
{
## Sets the Progress bar 100% value to total number of accounts in the import file
$progressBar.Maximum=$users.Count
     
## Converts the text password in CSV to an entry that can be stored in the AD in a secure form (eg hashed passwords, rather than plaintext)
## Then sets up the variabes used to create the AD account
$Password = (ConvertTo-SecureString -AsPlainText $user.Password -Force)
$Parameters = @{
       'Path'                 = $global:oupath
       'Name'                 = $user.Displayname
       'GivenName'            = $user.Firstname
       'Surname'              = $user.Lastname 
       'Displayname'          = $user.Displayname 
       'Description'          = $user.Description 
       'UserPrincipalName'    = $user.Username+’@said.oxford.edu’ 
       'SAMAccountName'       = $user.Username
       'ScriptPath'           = 'logonstub.exe'
       'Enabled'              = $true
       'AccountPassword'      = $Password
       'ChangePasswordAtLogon'= $true
       'PasswordNeverExpires' = $false
             }
## This section creates the accounts. If the AD Object does NOT already exist, create it, and add entry to the log file
try {
      
     if (!(get-aduser -Filter {SAMAccountName -eq "$user.Username"}))
     {
       New-ADUser @Parameters
       Add-ADGroupMember -Identity "Students_All" -Members $user.Username
       Add-ADGroupMember -Identity "SBS-UG-Require MFA for Users (on-prem)" -Members $user.Username
       Write-Verbose "[PASS] Created [$($user.Username)]"
       $global:SuccessLog += $user.Username + " PASS - AD Account created sucessfully"
       $global:s++
       $count++
       $progressBar.PerformStep()                                     
      }
    }
      
catch {
        Write-Warning "[WARNING] SAMAccountName for [$($user.Username)] already exists OR other error"
        $global:ErrorLog += $user.Username + " FAILED"
        $global:f++
        $count++
        $progressBar.PerformStep()                                
       }
[System.Windows.Forms.Application]::DoEvents()

## Writes the logs to disk
$ErrorLog | out-file -FilePath  $LogFolder\$classyear\fail-$date.log -Force
$SuccessLog | out-file -FilePath  $LogFolder\$classyear\success-$date.log -Force
}

Stop-Transcript
[System.Windows.Forms.MessageBox]::Show("Account creation complete!`n`nAccounts Sucessfully Created : $s`nAccounts Failed : $f  `n`nPlease review Event Logs if errors appear")
}
)

## Below function displays the File Selection popup box
Function Get-FileName($initialDirectory, $filter)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.Title = "Please select your Network Account batch .csv import file"
    $OpenFileDialog.filter = $filter
    #$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    #$OpenFileDialog.filter = "Log Files (*.log)| *.log"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$main_form.ShowDialog()