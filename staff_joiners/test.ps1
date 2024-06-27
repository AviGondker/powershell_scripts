# Staff Account Creation Script
# Create Staff network Accounts
# Avi Gondker 2024

#Setup Variables
$username = ""
$manager = ""
$title = ""
$site = ""
$location = ""
$department = ""

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text =’Staff Account Creation Script'
$form.Width = 1000
$form.Height = 700
$form.AutoSize = $false
$form.StartPosition = 'CenterScreen'
#$main_form.Topmost = $true

#Add Instructions Text
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Paste the JML info below"
$Label.Location  = New-Object System.Drawing.Point(10,30)
$Label.AutoSize = $true
$form.Controls.Add($Label)

#Add the Text Box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,50) 
$textBox.Size = New-Object System.Drawing.Size(600,450) ### Size of the text box
$textBox.Multiline = $true ### Allows multiple lines of data
$textBox.ScrollBars = "Vertical" ### Allows for a vertical scroll bar if the list of text is too big for the window
$form.Controls.Add($textBox)

#Add the Extract Data Button
$ExtractDataButton = New-Object System.Windows.Forms.Button
$ExtractDataButton.Location = New-Object System.Drawing.Point(700,50)
$ExtractDataButton.Size = New-Object System.Drawing.Size(100,40)
$ExtractDataButton.Text = "Extract Data..."
$form.Controls.Add($ExtractDataButton)

#Add Confirm Details label
$ConfirmDetailsLabel = New-Object System.Windows.Forms.Label
$ConfirmDetailsLabel.Text = "The data below will be used to create the account`n" + "Please CHECK then click OK"
$ConfirmDetailsLabel.Location  = New-Object System.Drawing.Point(630,100)
$ConfirmDetailsLabel.AutoSize = $true
$form.Controls.Add($ConfirmDetailsLabel)


## Add the CREATE Button
$createButton = New-Object System.Windows.Forms.Button
$createButton.Location = New-Object System.Drawing.Point(660,270)
$createButton.Size = New-Object System.Drawing.Size(160,40)
$createButton.Text = 'Create...!'
$form.Controls.Add($createButton)

##  Add the CLOSE Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(660,350)
$exitButton.Size = New-Object System.Drawing.Size(160,40)
$exitButton.Text = 'EXIT'
$exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $exitButton
$form.Controls.Add($exitButton)

#What Happens when you click the Extract Button
$ExtractDataButton.Add_Click(
{
$lines = $textBox.Text -split ":"


$username = $variables[2]
#$manager = $variables[5]
#$title = $variables[11]
#$site = $Variables[15]
#$location = $variables[17]
$department = $variables['Variable3']


#Add Username field
$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username is : $username"
$usernameLabel.Location  = New-Object System.Drawing.Point(630,140)
$usernameLabel.AutoSize = $true
$form.Controls.Add($usernameLabel)

#Add Manager field
$managerLabel = New-Object System.Windows.Forms.Label
$managerLabel.Text = "manager is : $manager"
$managerLabel.Location  = New-Object System.Drawing.Point(630,160)
$managerLabel.AutoSize = $true
$form.Controls.Add($managerLabel)

#Add Jobtitle Field
$jobTitleLabel = New-Object System.Windows.Forms.Label
$jobTitleLabel.Text = "Job Title is : $title"
$jobTitleLabel.Location  = New-Object System.Drawing.Point(630,180)
$jobTitleLabel.AutoSize = $true
$form.Controls.Add($jobTitleLabel )

#Add the Site field
$siteLabel = New-Object System.Windows.Forms.Label
$siteLabel.Text = "Site is : $site"
$siteLabel.Location  = New-Object System.Drawing.Point(630,200)
$siteLabel.AutoSize = $true
$form.Controls.Add($siteLabel)

#Add Office field
$officeLabel = New-Object System.Windows.Forms.Label
$officeLabel.Text = "Office is : $office"
$officeLabel.Location  = New-Object System.Drawing.Point(630,220)
$officeLabel.AutoSize = $true
$form.Controls.Add($officeLabel)

#Add Department field
$departmentLabel = New-Object System.Windows.Forms.Label
$departmentLabel.Text = "Department is : $department"
$departmentLabel.Location  = New-Object System.Drawing.Point(630,240)
$departmentLabel.AutoSize = $true
$form.Controls.Add($departmentLabel)
}
)

$form.ShowDialog()