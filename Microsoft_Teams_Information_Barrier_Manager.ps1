<# 
.Synopsis 
The purpose of this tool is to give you an easy front end for the management of Microsoft Teams Information Barriers.

.DESCRIPTION 
PowerShell GUI script which allows for GUI management of Teams Information Barriers

.Notes 
NAME:      Microsoft_Teams_Information_Barrier_Manager.ps1
VERSION:   1.0
AUTHOR:    C. Anthony Caragol 
LASTEDIT:  02/25/2020

V 1.0 - February 02, 2020 - Initial release 

.Link 
Website: http://www.teamsadmin.com
Twitter: http://www.twitter.com/canthonycaragol
LinkedIn: http://www.linkedin.com/in/canthonycaragol

.EXAMPLE 
.\Microsoft_Teams_Information_Barrier_Manager.ps1

.TODO
Maybe autotest prereqs?  Who's not licensed?
Test to see where symmetry is missing?

.APOLOGY
Please excuse the sloppy coding, I don't use a development environment, IDE or ISE.  I use notepad, 
not even Notepad++, just notepad.  I am not a developer, just an enthusiast so some code may be redundant or
inefficient.
#>


$Global:TeamsAdminIcon = [System.Convert]::FromBase64String('
AAABAAEAJiEAAAEACABYCgAAFgAAACgAAAAmAAAAQgAAAAEACAAAAAAAKAUAAAAAAAAAAAAAAAEAAAAB
AAAAAAAAMwAAAGYAAACZAAAAzAAAAP8AAAAAKwAAMysAAGYrAACZKwAAzCsAAP8rAAAAVQAAM1UAAGZV
AACZVQAAzFUAAP9VAAAAgAAAM4AAAGaAAACZgAAAzIAAAP+AAAAAqgAAM6oAAGaqAACZqgAAzKoAAP+q
AAAA1QAAM9UAAGbVAACZ1QAAzNUAAP/VAAAA/wAAM/8AAGb/AACZ/wAAzP8AAP//AAAAADMAMwAzAGYA
MwCZADMAzAAzAP8AMwAAKzMAMyszAGYrMwCZKzMAzCszAP8rMwAAVTMAM1UzAGZVMwCZVTMAzFUzAP9V
MwAAgDMAM4AzAGaAMwCZgDMAzIAzAP+AMwAAqjMAM6ozAGaqMwCZqjMAzKozAP+qMwAA1TMAM9UzAGbV
MwCZ1TMAzNUzAP/VMwAA/zMAM/8zAGb/MwCZ/zMAzP8zAP//MwAAAGYAMwBmAGYAZgCZAGYAzABmAP8A
ZgAAK2YAMytmAGYrZgCZK2YAzCtmAP8rZgAAVWYAM1VmAGZVZgCZVWYAzFVmAP9VZgAAgGYAM4BmAGaA
ZgCZgGYAzIBmAP+AZgAAqmYAM6pmAGaqZgCZqmYAzKpmAP+qZgAA1WYAM9VmAGbVZgCZ1WYAzNVmAP/V
ZgAA/2YAM/9mAGb/ZgCZ/2YAzP9mAP//ZgAAAJkAMwCZAGYAmQCZAJkAzACZAP8AmQAAK5kAMyuZAGYr
mQCZK5kAzCuZAP8rmQAAVZkAM1WZAGZVmQCZVZkAzFWZAP9VmQAAgJkAM4CZAGaAmQCZgJkAzICZAP+A
mQAAqpkAM6qZAGaqmQCZqpkAzKqZAP+qmQAA1ZkAM9WZAGbVmQCZ1ZkAzNWZAP/VmQAA/5kAM/+ZAGb/
mQCZ/5kAzP+ZAP//mQAAAMwAMwDMAGYAzACZAMwAzADMAP8AzAAAK8wAMyvMAGYrzACZK8wAzCvMAP8r
zAAAVcwAM1XMAGZVzACZVcwAzFXMAP9VzAAAgMwAM4DMAGaAzACZgMwAzIDMAP+AzAAAqswAM6rMAGaq
zACZqswAzKrMAP+qzAAA1cwAM9XMAGbVzACZ1cwAzNXMAP/VzAAA/8wAM//MAGb/zACZ/8wAzP/MAP//
zAAAAP8AMwD/AGYA/wCZAP8AzAD/AP8A/wAAK/8AMyv/AGYr/wCZK/8AzCv/AP8r/wAAVf8AM1X/AGZV
/wCZVf8AzFX/AP9V/wAAgP8AM4D/AGaA/wCZgP8AzID/AP+A/wAAqv8AM6r/AGaq/wCZqv8AzKr/AP+q
/wAA1f8AM9X/AGbV/wCZ1f8AzNX/AP/V/wAA//8AM///AGb//wCZ//8AzP//AP///wAAAAAAAAAAAAAA
AAAAAAAAHB0WHRwdHRwXHRwdHRwdHB0cFx0cHRwXHRwdHRwdHB0dFh0dHB0AAB0cHRwXHB0WHRwXHB0W
HRYdHB0cFxwXHB0WHRwXHBccHRwdFh0cAAAcHRYdHB0cHRwdHB0cHRwdHB0WHRwdHB0cHRwdHB0cHRwX
HB0cHQAAHRwdHBcdFh0cFxwdFh0XHB0dHB0WHRwXHBccFxwdFh0cHRwXHB0AAB0cFxwdHB0cHR0cHR0c
HRwdHBccHR0cHR0cHR0cHR0cHRccHRwdAAAdHB0dFh0cFxwXHBccFxwdFh0cHRYdFh0cFxwdFh0WHRwd
HBccHQAAHRwXHB0cHRwdHB0cHRwdHB0WHRwdHB0cHRwdHB0cHRwdFh0dHB0AAB0cHRwXHBccHRwXHB0c
Fx0cHRwXHB0cFxwXHRYdHBccHRwdFh0cAAAdHBccHRwdHBccHRwXHB0cHRYdHB0WHRwdHB0cHRwdHRwX
HB0cHQAAHRwdd/v7+/v7+/v7+/v7mh3R+/v7+/v1HBccHRYdHB0cHRwXHB0AAB0WHSL7+/v7+/v7+/v7
+/UcTfv7+/v7+3AdHB0dFh0WHRccHRwdAAAdHB0X0fv7+/v7+/v7+/v7mh3R+/v7+/uaHRYdHB0cHRwd
HBccHQAAFh0cHU37+/v7+/v7+/v7+/Udp/v7+/v7yxwdHB0cFxwdFh0dFh0AAB0cFxwd0fv7+/v7+/v7
+/v7cB37+/v7+/tAHRYdHB0dHB0cHRwdAAAdHB0cF6f7+/v7+8oXHB0cHR0c0fv7+/v7+/v7+/vEHRwd
Fh0cHQAAHRYdHRxN+/v7+/v1HB0XHB0WHXf7+/v7+/v7+/v79BccFxwdFxwAAB0cHRwXHNH7+/v7+3Ad
HB0cHRwd+/v7+/v7+/v7+/tHHB0cHRwdAAAdFh0cHR2n+/v7+/ubHB0WHRwdHdH7+/v7+/v7+/v7xB0c
HRYdHAAAHRwdHB0cTfv7+/v79B0cHRwdFh13+/v7+/v7+/v7+/UdFh0dHB0AAB0cHRccFx37+/v7+/tw
HRwXHB0cHfv7+/v7+0YdHB0cHRwcHRwXAAAcFxwdHB0c0fv7+/v7xRYdHRwdHB3R+/v7+/v7+/v7+/vF
HRYdHAAAHRwdFh0cHXf7+/v7+/tHHB0WHRccd/v7+/v7+/v7+/v79B0cHR0AABccHRwdFh0d+/v7+/v7
mhccHRwdHE37+/v7+/v7+/v7+/UdHBccAAAdHBcdHB0cHdH7+/v7+8scHRYdHB0d0fv7+/v7+/v7+/v7
cB0cHQAAHB0cHRwdFh13+/v7+/v7ah0dHB0WHaH7+/v7+/v7+/v7+8UcHR0AAB0cFxwXHB0cHRwdFh0c
HRwdHBccHRwdHB0WHRwXHB0cHRwXHBccAAAWHRwdHB0dHB0WHRwdHRwdFh0cHRwdHRYdHB0cHR0WHRYd
HB0cHQAAHRwXHB0cFxwXHB0WHRYdFh0cFxwXHBccHRYdFh0cHRwdHBccHRwAAB0cHR0WHRwdHRwdHRwd
HB0dHB0dHB0cHR0cHR0cHRwdFxwdFxwdAAAdFh0cHRwXHB0WHRwdFh0cFxwdHBccHRYdHB0WHRwXHB0c
HRwdHAAAHB0cHRYdHB0cHRwXHB0cHRwdFh0cHRwdHBccHRwdHB0WHRwdFh0AAB0cFxwdHBccHRYdHB0W
HRwXHB0cFxwdFh0cHRYdHBccHRwXHB0cAAAcHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0c
HR0cHQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
')



Function ConnectToTenant()
{
	$UserCredential = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session -DisableNameChecking
	RefreshMainFormData
}

Function RefreshMainFormData()
{
    $OrgSegmentDatagrid.Rows.Clear()
	$Global:AllSegments=@()
        foreach ($OrgSeg in (Get-OrganizationSegment)) { $Global:AllSegments += $OrgSeg.name;$OrgSegmentDatagrid.Rows.Add($OrgSeg.name,$OrgSeg.UserGroupFilter) }

    $IBPolicyDatagrid.Rows.Clear()
        foreach ($Policy in (Get-InformationBarrierPolicy)) 
		{
			if ($Policy.SegmentsBlocked -eq $null) {
				$SegmentsAllowed=$Policy.SegmentsAllowed -join ', '    
				$IBPolicyDatagrid.Rows.Add($Policy.name,$Policy.AssignedSegment,"Allowed",$Policy.SegmentsAllowed,$Policy.Comment) 
			}
			else
			{
				$SegmentsBlocked=$Policy.SegmentsBlocked -join ', '    
				$IBPolicyDatagrid.Rows.Add($Policy.name,$Policy.AssignedSegment,"Blocked",$SegmentsBlocked,$Policy.Comment) 
			}
		}
}

Function ShowHelp()
{
	$ShowHelpForm = New-Object System.Windows.Forms.Form 
	$ShowHelpForm.Text = "Why won't my barriers work?!?!!"
	$ShowHelpForm.Size = New-Object System.Drawing.Size(600,400) 
	$ShowHelpForm.MinimumSize = New-Object System.Drawing.Size(600,400) 
	$ShowHelpForm.StartPosition = "CenterScreen"
	$ShowHelpForm.KeyPreview = $True
	$ShowHelpForm.Icon = $Global:TeamsAdminIcon
    
	$HelpRichTextbox = New-Object System.Windows.Forms.label
	$HelpRichTextbox.Location = New-Object System.Drawing.Size(30,10) 
	$HelpRichTextbox.Size = New-Object System.Drawing.Size(500,250) 
	$HelpRichTextbox.Text = "Implementing Barriers unfortunately isn't as simple as creating segments and policies then clicking Apply. There are prerequisties that need to be in place, and some rules to follow.  I will highlight some of the important items I see missed here. `r`n`r`n`r`n   1)  All users that the information barrier protects needs to be licensed with E5 or Advanced Compliance. `r`n`r`n   2)  As an administrator, you need to belong to at least one of the following roles: Global Administrator, Compliance Administrator, IB Compliance Management `r`n`r`n   3)  There should be no Exchange Address Policies in place. `r`n`r`n   4)  Scope directory search needs to be enabled in Microsoft Teams. `r`n`r`n    5)  Policies need to be symmetric.  Basically, if Sales can't talk to Compliance, then you need a second policy so that Compliance can't talk to Sales. `r`n`r`n    6)  Click the link below for more prerequisites."
	$ShowHelpForm.Controls.Add($HelpRichTextbox)
    
	$IBPrereqsLinkLabel = New-Object System.Windows.Forms.LinkLabel
	$IBPrereqsLinkLabel.Location = New-Object System.Drawing.Size(30,270) 
	$IBPrereqsLinkLabel.Size = New-Object System.Drawing.Size(500,40)
	$IBPrereqsLinkLabel.text = "https://docs.microsoft.com/en-us/microsoft-365/compliance/information-barriers-policies?view=o365-worldwide#prerequisites"
	$IBPrereqsLinkLabel.add_Click({Start-Process $IBPrereqsLinkLabel.text})
	$IBPrereqsLinkLabel.Anchor = 'Bottom, Left'
	$ShowHelpForm.Controls.Add($IBPrereqsLinkLabel)
    
	$GotItButton = New-Object System.Windows.Forms.Button
	$GotItButton.Location = New-Object System.Drawing.Size(30,300)
	$GotItButton.Size = New-Object System.Drawing.Size(500,40)
	$GotItButton.Text = "Got It"
	$GotItButton.Add_Click(
	{
		$ShowHelpForm.Close()
	})
	$GotItButton.Anchor = 'Bottom, Left'
	$ShowHelpForm.Controls.Add($GotItButton)
   
	[void] $ShowHelpForm.ShowDialog()
}

Function EditSegment()
{
	param( [string]$CreateOrEdit, [string]$SegmentName, [string]$SegmentFilter )
	$EditSegmentForm = New-Object System.Windows.Forms.Form 
	if ($CreateOrEdit -eq "Edit") {$EditSegmentForm.Text = "Edit Segment - $SegmentName"} else {$EditSegmentForm.Text = "Create New Segment"}
	$EditSegmentForm.Size = New-Object System.Drawing.Size(600,240) 
	$EditSegmentForm.MinimumSize = New-Object System.Drawing.Size(600,240) 
	$EditSegmentForm.StartPosition = "CenterScreen"
	$EditSegmentForm.KeyPreview = $True
	$EditSegmentForm.Icon = $Global:TeamsAdminIcon
    
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size(30,10) 
	$NameLabel.Size = New-Object System.Drawing.Size(200,20) 
	$NameLabel.Text = "Name:"
	$EditSegmentForm.Controls.Add($NameLabel)
    
	$NameTextbox = New-Object System.Windows.Forms.Textbox
	$NameTextbox.Location = New-Object System.Drawing.Size(30,30) 
	$NameTextbox.Size = New-Object System.Drawing.Size(200,20) 
	if ($CreateOrEdit -eq "Edit") {$NameTextbox.Enabled=$False} else {$NameTextbox.Enabled=$True}
	$NameTextbox.Text = $SegmentName
	$EditSegmentForm.Controls.Add($NameTextbox) 
    
	$FilterLabel = New-Object System.Windows.Forms.Label
	$FilterLabel.Location = New-Object System.Drawing.Size(30,55) 
	$FilterLabel.Size = New-Object System.Drawing.Size(200,20) 
	$FilterLabel.Text = "Filter:"
	$EditSegmentForm.Controls.Add($FilterLabel)
    
	$FilterTextbox = New-Object System.Windows.Forms.Textbox
	$FilterTextbox.Location = New-Object System.Drawing.Size(30,75) 
	$FilterTextbox.Size = New-Object System.Drawing.Size(200,80) 
	$FilterTextbox.Text = $SegmentFilter
	$FilterTextbox.Multiline = $True
	$EditSegmentForm.Controls.Add($FilterTextbox) 
    
	$FilterResultLabel = New-Object System.Windows.Forms.Label
	$FilterResultLabel.Location = New-Object System.Drawing.Size(240,30) 
	$FilterResultLabel.Size = New-Object System.Drawing.Size(300,150) 
	$FilterResultLabel.Text = "Filter Examples: `nDepartment -eq 'Sales'`r`n`r`nDepartment -ne 'Human Resources'`r`n`r`n(Description -eq ""Salaried"") -and (City -eq ""Chicago"")`r`n`r`n(Description -eq ""Salaried"") -and ((City -eq ""Chicago"") -or (City -eq ""Des Moines""))"
	$EditSegmentForm.Controls.Add($FilterResultLabel)
    
	$SaveSegmentChangesButton = New-Object System.Windows.Forms.Button
	$SaveSegmentChangesButton.Location = New-Object System.Drawing.Size(30,160)
	$SaveSegmentChangesButton.Size = New-Object System.Drawing.Size(100,25)
	if ($CreateOrEdit -eq "Edit") {$SaveSegmentChangesButton.Text = "Save Edit"} else {$SaveSegmentChangesButton.Text = "Create"}
	$SaveSegmentChangesButton.Add_Click(
	{
		$SaveSegmentChangesButton.enabled=$False
		$CancelButton.enabled=$False
		$SaveSegmentChangesButton.text="Saving..."
	
		if ($CreateOrEdit -eq "Edit") 
		{
			$errorcount=$error.count
			Get-OrganizationSegment |Where {$_.Name -like $SegmentName}|Set-OrganizationSegment -UserGroupFilter $FilterTextbox.Text
			if ($error.count -gt $errorcount) 
			{
				[Microsoft.VisualBasic.Interaction]::MsgBox("Error: $(($error[0] -split "Status:")[0])" ,'Exclamation', "Error.")
			}
			RefreshMainFormData  
			$SaveSegmentChangesButton.enabled=$true
			$CancelButton.enabled=$true
			$SaveSegmentChangesButton.text="Save Edit"
			$EditSegmentForm.Close()   
		} 
		else 
		{
			$errorcount=$error.count
			New-OrganizationSegment -Name $NameTextbox.Text -UserGroupFilter $FilterTextbox.Text
			if ($error.count -gt $errorcount) 
			{
				[Microsoft.VisualBasic.Interaction]::MsgBox("Error: $(($error[0] -split "Status:")[0])" ,'Exclamation', "Error.")
			}
			RefreshMainFormData 
			$SaveSegmentChangesButton.enabled=$true
			$CancelButton.enabled=$true
			$SaveSegmentChangesButton.text="Create"
			$EditSegmentForm.Close()    
		}

	})
	$SaveSegmentChangesButton.Anchor = 'Bottom, Left'
	$EditSegmentForm.Controls.Add($SaveSegmentChangesButton)
    
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(130,160)
	$CancelButton.Size = New-Object System.Drawing.Size(100,25)
	$CancelButton.Text = "Quit"
	$CancelButton.Add_Click({
		$EditSegmentForm.Close()
	})
	$CancelButton.Anchor = 'Bottom, Left'
	$EditSegmentForm.Controls.Add($CancelButton)
    
	[void] $EditSegmentForm.ShowDialog()
}

Function SegmentSelection()
{
	param( [string]$PassedSegments )
	$returnvalue=""
	$SegmentArray=($PassedSegments.split(',')| % { $_.Trim() })
	$firsttime = $true

	$SegmentSelectionForm = New-Object System.Windows.Forms.Form 
	$SegmentSelectionForm.Text = "Select Segments"
	$SegmentSelectionForm.Size = New-Object System.Drawing.Size(300,330) 
	$SegmentSelectionForm.MinimumSize = New-Object System.Drawing.Size(300,330) 
	$SegmentSelectionForm.StartPosition = "CenterScreen"
	$SegmentSelectionForm.KeyPreview = $True
	$SegmentSelectionForm.Icon = $Global:TeamsAdminIcon
	$SegmentSelectionForm.add_Shown({
    		for($i=0;$i -lt $NameTextbox.items.Count;$i++) {
			if ($SegmentArray -contains $NameTextbox.items[$i].tostring().trim()) {$NameTextbox.setitemchecked($i,$true)}
        	}
	})
    
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size(10,10) 
	$NameLabel.Size = New-Object System.Drawing.Size(200,20) 
	$NameLabel.Text = "Select Segements"
	$SegmentSelectionForm.Controls.Add($NameLabel)
    
	$NameTextbox = New-Object System.Windows.Forms.CheckedListBox
	$NameTextbox.Location = New-Object System.Drawing.Size(10,30) 
	$NameTextbox.Size = New-Object System.Drawing.Size(250,200) 
	$NameTextbox.DataSource = [collections.arraylist]$Global:AllSegments
	$NameTextbox.CheckOnClick=$true
	$SegmentSelectionForm.Controls.Add($NameTextbox) 
 
	$SaveSegmentChangesButton = New-Object System.Windows.Forms.Button
	$SaveSegmentChangesButton.Location = New-Object System.Drawing.Size(10,250)
	$SaveSegmentChangesButton.Size = New-Object System.Drawing.Size(125,25)
	$SaveSegmentChangesButton.Text = "Save Edit"
	$SaveSegmentChangesButton.Add_Click({
		$returnvalue=$NameTextbox.checkeditems
		$SegmentSelectionForm.Close()
   	})
	$SaveSegmentChangesButton.Anchor = 'Bottom, Left'
	$SegmentSelectionForm.Controls.Add($SaveSegmentChangesButton)
    
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(135,250)
	$CancelButton.Size = New-Object System.Drawing.Size(125,25)
	$CancelButton.Text = "Quit"
	$CancelButton.Add_Click({
		for($i=0;$i -lt $NameTextbox.items.Count;$i++)
		{ 
			$NameTextbox.SetItemChecked($i, $false)
		}
		$SegmentSelectionForm.Close()
	})
	$CancelButton.Anchor = 'Bottom, Left'
	$SegmentSelectionForm.Controls.Add($CancelButton)

	[void] $SegmentSelectionForm.ShowDialog()
	return $NameTextbox.checkeditems
}

Function EditPolicy()
{
	param( [string]$CreateOrEdit, [string]$PolicyName, [string]$PolicyAssignment, [string]$PolicyAllowBlock, [string]$PolicySegments, [string]$PolicyComment )
	$EditPolicyForm = New-Object System.Windows.Forms.Form 
	if ($CreateOrEdit -eq "Edit") {$EditPolicyForm.Text = "Edit Policy - $PolicyName"} else {$EditPolicyForm.Text = "Create New Policy"}
	$EditPolicyForm.Size = New-Object System.Drawing.Size(380,400) 
	$EditPolicyForm.MinimumSize = New-Object System.Drawing.Size(380,400) 
	$EditPolicyForm.StartPosition = "CenterScreen"
	$EditPolicyForm.KeyPreview = $True
	$EditPolicyForm.Icon = $Global:TeamsAdminIcon
    
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size(30,10) 
	$NameLabel.Size = New-Object System.Drawing.Size(200,20) 
	$NameLabel.Text = "Name:"
	$EditPolicyForm.Controls.Add($NameLabel)
    
	$NameTextbox = New-Object System.Windows.Forms.Textbox
	$NameTextbox.Location = New-Object System.Drawing.Size(30,30) 
	$NameTextbox.Size = New-Object System.Drawing.Size(300,20) 
	$NameTextbox.Text = $PolicyName
	if ($CreateOrEdit -eq "Edit") {$NameTextbox.enabled = $false}
	$EditPolicyForm.Controls.Add($NameTextbox) 
    
	$PolicyAssignedLabel = New-Object System.Windows.Forms.Label
	$PolicyAssignedLabel.Location = New-Object System.Drawing.Size(30,55) 
	$PolicyAssignedLabel.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyAssignedLabel.Text = "Assigned To:"
	$EditPolicyForm.Controls.Add($PolicyAssignedLabel)
    
	$PolicyAssignedComboBox = New-Object System.Windows.Forms.Combobox
	$PolicyAssignedComboBox.Location = New-Object System.Drawing.Size(30,75) 
	$PolicyAssignedComboBox.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyAssignedComboBox.Text = $PolicyAssignment
	for($i=0;$i -lt $OrgSegmentDatagrid.RowCount;$i++)
	{ 
		$PolicyAssignedComboBox.Items.add($OrgSegmentDatagrid.Rows[$i].Cells[0].Value)
	}
	if ($CreateOrEdit -eq "Edit") {$PolicyAssignedComboBox.enabled = $false}
	$EditPolicyForm.Controls.Add($PolicyAssignedComboBox) 
    
	$PolicyAllowedBlockedLabel = New-Object System.Windows.Forms.Label
	$PolicyAllowedBlockedLabel.Location = New-Object System.Drawing.Size(30,100) 
	$PolicyAllowedBlockedLabel.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyAllowedBlockedLabel.Text = "Allowed or Blocked:"
	$EditPolicyForm.Controls.Add($PolicyAllowedBlockedLabel)
    
	$PolicyAllowedBlockedTextbox = New-Object System.Windows.Forms.Combobox
	$PolicyAllowedBlockedTextbox.Location = New-Object System.Drawing.Size(30,120) 
	$PolicyAllowedBlockedTextbox.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyAllowedBlockedTextbox.Text = $PolicyAllowBlock
	$PolicyAllowedBlockedTextbox.Items.Add("Allowed")
	$PolicyAllowedBlockedTextbox.Items.Add("Blocked")
	$PolicyAllowedBlockedTextbox.Text="Blocked"
	if ($CreateOrEdit -eq "Edit") {$PolicyAllowedBlockedTextbox.enabled = $false}
	$EditPolicyForm.Controls.Add($PolicyAllowedBlockedTextbox) 
    
	$PolicySegmentsLabel = New-Object System.Windows.Forms.Label
	$PolicySegmentsLabel.Location = New-Object System.Drawing.Size(30,145) 
	$PolicySegmentsLabel.Size = New-Object System.Drawing.Size(300,20) 
	$PolicySegmentsLabel.Text = "Policy Segments:"
	$EditPolicyForm.Controls.Add($PolicySegmentsLabel)
    
	$PolicySegmentsTextbox = New-Object System.Windows.Forms.Textbox
	$PolicySegmentsTextbox.Location = New-Object System.Drawing.Size(30,165) 
	$PolicySegmentsTextbox.Size = New-Object System.Drawing.Size(250,60) 
	$PolicySegmentsTextbox.Text = $PolicySegments
	$PolicySegmentsTextbox.Multiline = $True
	$EditPolicyForm.Controls.Add($PolicySegmentsTextbox) 
    
	$PolicySegmentsButton = New-Object System.Windows.Forms.Button
	$PolicySegmentsButton.Location = New-Object System.Drawing.Size(290,165) 
	$PolicySegmentsButton.Size = New-Object System.Drawing.Size(40,20) 
	$PolicySegmentsButton.Text = "Edit"
	$PolicySegmentsButton.Add_Click({
		$SelectedSegments=SegmentSelection $PolicySegmentsTextbox.Text
		if ($SelectedSegments.count -gt 0) {
			$PolicySegmentsTextbox.Text=""
			Foreach ($Segment in $SelectedSegments) 
			{
				if ($PolicySegmentsTextbox.Text.length -gt 0) 
				{
					$PolicySegmentsTextbox.Text +=", $Segment"
				}
				else
				{
					$PolicySegmentsTextbox.Text = $Segment
				}	
			}
		}
	})
	$EditPolicyForm.Controls.Add($PolicySegmentsButton) 

	$PolicyCommentLabel = New-Object System.Windows.Forms.Label
	$PolicyCommentLabel.Location = New-Object System.Drawing.Size(30,230) 
	$PolicyCommentLabel.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyCommentLabel.Text = "Comments:"
	$EditPolicyForm.Controls.Add($PolicyCommentLabel)
    
	$PolicyCommentTextbox = New-Object System.Windows.Forms.Textbox
	$PolicyCommentTextbox.Location = New-Object System.Drawing.Size(30,250) 
	$PolicyCommentTextbox.Size = New-Object System.Drawing.Size(300,40) 
	$PolicyCommentTextbox.Text = $PolicyComment
	$PolicyCommentTextbox.Multiline = $True
	$EditPolicyForm.Controls.Add($PolicyCommentTextbox) 
    
	$SavePolicyChangesButton = New-Object System.Windows.Forms.Button
	$SavePolicyChangesButton.Location = New-Object System.Drawing.Size(30,300)
	$SavePolicyChangesButton.Size = New-Object System.Drawing.Size(150,25)
	if ($CreateOrEdit -eq "Edit") {$SavePolicyChangesButton.Text = "Save Edit"} else {$SavePolicyChangesButton.Text = "Create"}
	$SavePolicyChangesButton.Add_Click({
		$SavePolicyChangesButton.enabled=$False
		$CancelButton.enabled=$False
		$SavePolicyChangesButton.text="Saving..."
	
		$newsegmentselection=@()
		foreach ($d in ($PolicySegmentsTextbox.Text.split(',')| % { $_.Trim() })) {$newsegmentselection += $d}
		if ($CreateOrEdit -eq "Edit") 
		{
			$errorcount=$error.count

			$OldPolicy = Get-InformationBarrierPolicy | Where {$_.Name -like $NameTextbox.Text}
			if ($PolicyAllowedBlockedTextbox.Text = "Blocked") 
			{
				Set-InformationBarrierPolicy -identity $OldPolicy.GUID -SegmentsBlocked $newsegmentselection -State Active -Comment $PolicyCommentTextbox.Text
			}
			else
			{
				Set-InformationBarrierPolicy -identity $OldPolicy.GUID -SegmentsAllowed $newsegmentselection -State Active -Comment $PolicyCommentTextbox.Text
			}
			if ($error.count -gt $errorcount) 
			{
				[Microsoft.VisualBasic.Interaction]::MsgBox("Error: $(($error[0] -split "Status:")[0])" ,'Exclamation', "Error.")
			}
			RefreshMainFormData  
			$SavePolicyChangesButton.enabled=$true
			$CancelButton.enabled=$true
			$SavePolicyChangesButton.text="Save Edit"
			$EditPolicyForm.Close()   
		} 
		else 
		{
			$errorcount=$error.count
			if ($PolicyAllowedBlockedTextbox.Text = "Blocked") 
				{
					New-InformationBarrierPolicy -Name $NameTextbox.Text -AssignedSegment $PolicyAssignedComboBox.Text -SegmentsBlocked $newsegmentselection -State Active -Comment $PolicyCommentTextbox.Text
				}
			else
				{
					New-InformationBarrierPolicy -Name $NameTextbox.Text -AssignedSegment $PolicyAssignedComboBox.Text -SegmentsAllowed $newsegmentselection -State Active -Comment $PolicyCommentTextbox.Text
				}

			if ($error.count -gt $errorcount) 
			{
				[Microsoft.VisualBasic.Interaction]::MsgBox("Error: $(($error[0] -split "Status:")[0])" ,'Exclamation', "Error.")
			}
			RefreshMainFormData 
			$SavePolicyChangesButton.enabled=$true
			$CancelButton.enabled=$true
			$SavePolicyChangesButton.text="Create"
			$EditPolicyForm.Close()    
		}
	})
	$SavePolicyChangesButton.Anchor = 'Bottom, Left'
	$EditPolicyForm.Controls.Add($SavePolicyChangesButton)
    
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(190,300)
	$CancelButton.Size = New-Object System.Drawing.Size(150,25)
	$CancelButton.Text = "Quit"
	$CancelButton.Add_Click({
		$EditPolicyForm.Close()
	})
	$CancelButton.Anchor = 'Bottom, Left'
	$EditPolicyForm.Controls.Add($CancelButton)
    
	[void] $EditPolicyForm.ShowDialog()
}

Function MainForm()
{
    
	$mainForm = New-Object System.Windows.Forms.Form 
	$mainForm.Text = "Microsoft Teams Information Barrier Manager v 1.0.0"
	$mainForm.Size = New-Object System.Drawing.Size(1000,560) 
	$mainForm.MinimumSize = New-Object System.Drawing.Size(1000,560) 
	$mainForm.StartPosition = "CenterScreen"
	$mainForm.Add_SizeChanged($CAC_FormSizeChanged)
	$mainForm.KeyPreview = $True
	$mainForm.Icon = $Global:TeamsAdminIcon
    
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size(10,10) 
	$TitleLabel.Size = New-Object System.Drawing.Size(780,40) 
	$TitleLabel.Text = "The purpose of this tool is to give you an easy graphical method of editing your Microsoft Teams Information Barrier settings in your tenant without having to remember a bunch of PowerShell commands you'd rarely use.  Of course, applying them can be complex because some of the requirements can be complex.  I try to surface errors where I can, press the Help button for tips.  This is an early version with revisions coming soon."
	$mainForm.Controls.Add($TitleLabel) 
    
	$OrgSegmentDataGridLabel = New-Object System.Windows.Forms.Label
	$OrgSegmentDataGridLabel.Location = New-Object System.Drawing.Size(10,65) 
	$OrgSegmentDataGridLabel.Size = New-Object System.Drawing.Size(250,15) 
	$OrgSegmentDataGridLabel.Text = "Organization Segments"
	$mainForm.Controls.Add($OrgSegmentDataGridLabel) 
    
	$OrgSegmentDatagrid = New-Object System.Windows.Forms.DataGridView
	$OrgSegmentDatagrid.Location = New-Object System.Drawing.Size(10,80) 
	$OrgSegmentDatagrid.Size = New-Object System.Drawing.Size(345,365) 
	$OrgSegmentDatagrid.Anchor = 'Top, Bottom,Left'
	$OrgSegmentDatagrid.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Bisque 
	$OrgSegmentDatagrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Beige
	$OrgSegmentDatagrid.ColumnCount = 2
	$OrgSegmentDatagrid.SelectionMode = 'FullRowSelect'
	$OrgSegmentDatagrid.RowHeadersVisible = $False
	$OrgSegmentDatagrid.AllowUsertoAddRows = $False
	$OrgSegmentDatagrid.Columns[0].Width = 142
	$OrgSegmentDatagrid.Columns[1].Width = 200
	$OrgSegmentDatagrid.Columns[0].Name = "Name"
	$OrgSegmentDatagrid.Columns[1].Name = "Filter"
    
	$OrgSegmentDatagrid.Add_DoubleClick({
		$SelectedSegment= $OrgSegmentDatagrid.Rows[$OrgSegmentDatagrid.CurrentRow.Index].Cells[0].value
		$SelectedFilter= $OrgSegmentDatagrid.Rows[$OrgSegmentDatagrid.CurrentRow.Index].Cells[1].value
		EditSegment "Edit" $SelectedSegment $SelectedFilter
	})
	$mainForm.Controls.Add($OrgSegmentDatagrid) 

	$IBPolicyDataGridLabel = New-Object System.Windows.Forms.Label
	$IBPolicyDataGridLabel.Location = New-Object System.Drawing.Size(370,65) 
	$IBPolicyDataGridLabel.Size = New-Object System.Drawing.Size(250,15) 
	$IBPolicyDataGridLabel.Text = "Information Barrier Policies"
	$mainForm.Controls.Add($IBPolicyDataGridLabel) 

	$IBPolicyDatagrid = New-Object System.Windows.Forms.DataGridView
	$IBPolicyDatagrid.Location = New-Object System.Drawing.Size(370,80) 
	$IBPolicyDatagrid.Size = New-Object System.Drawing.Size(590,365) 
	$IBPolicyDatagrid.Anchor = 'Top, Bottom,Left'
	$IBPolicyDatagrid.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Bisque 
	$IBPolicyDatagrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Beige
	$IBPolicyDatagrid.ColumnCount = 5
	$IBPolicyDatagrid.Columns[1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
	$IBPolicyDatagrid.SelectionMode = 'FullRowSelect'
	$IBPolicyDatagrid.RowHeadersVisible = $False
	$IBPolicyDatagrid.AllowUsertoAddRows = $False
	$IBPolicyDatagrid.Columns[0].Name = "Name"
	$IBPolicyDatagrid.Columns[1].Name = "Assigned To"
	$IBPolicyDatagrid.Columns[2].Name = "Allowed\Blocked"
	$IBPolicyDatagrid.Columns[3].Name = "Segments"
	$IBPolicyDatagrid.Columns[4].Name = "Comment"
    
	$IBPolicyDatagrid.Add_DoubleClick(
	{
		$SelectedPolicy= $IBPolicyDatagrid.Rows[$IBPolicyDatagrid.CurrentRow.Index].Cells[0].value
		$SelectedFilter= $IBPolicyDatagrid.Rows[$IBPolicyDatagrid.CurrentRow.Index].Cells[1].value
		$SelectedAllowBlock= $IBPolicyDatagrid.Rows[$IBPolicyDatagrid.CurrentRow.Index].Cells[2].value
		$SelectedSegments= $IBPolicyDatagrid.Rows[$IBPolicyDatagrid.CurrentRow.Index].Cells[3].value
		$SelectedComments= $IBPolicyDatagrid.Rows[$IBPolicyDatagrid.CurrentRow.Index].Cells[4].value
		EditPolicy "Edit" $SelectedPolicy $SelectedFilter $SelectedAllowBlock $SelectedSegments $SelectedComments
	 })
	$mainForm.Controls.Add($IBPolicyDatagrid) 

	$MainFormButtonWidth=100
	$MainFormButtonHeight=40
	$MainFormButtonSpacing=6
	
	$ConnectTenantButton = New-Object System.Windows.Forms.Button
	$ConnectTenantButton.Location = New-Object System.Drawing.Size(((10 + (0 * $MainFormButtonWidth) + (0 * $MainFormButtonSpacing))  + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$ConnectTenantButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$ConnectTenantButton.Text = "Connect to Tenant"
	$ConnectTenantButton.Add_Click({
		$ConnectTenantButton.Enabled = $false
		$RefreshButton.Enabled = $false
		$CreateSegmentButton.Enabled = $false
		$CreatePolicyButton.Enabled = $false
		$StartApplicationButton.Enabled = $false
		$DeleteSegmentButton.Enabled=$false
		$DeletePolicyButton.Enabled = $false
		$ConnectTenantButton.Text = "Connecting..."
  	      ConnectToTenant
		$ConnectTenantButton.Enabled = $true
		$RefreshButton.Enabled = $true
		$CreateSegmentButton.Enabled = $true
		$CreatePolicyButton.Enabled = $true
		$DeleteSegmentButton.Enabled=$true
		$DeletePolicyButton.Enabled = $true
		$StartApplicationButton.Enabled = $true
		$ConnectTenantButton.Text = "Connect to Tenant"
	})
	$ConnectTenantButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($ConnectTenantButton)
   
	$RefreshButton = New-Object System.Windows.Forms.Button
	$RefreshButton.Location = New-Object System.Drawing.Size(((10 + (1 * $MainFormButtonWidth) + (1 * $MainFormButtonSpacing))  + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$RefreshButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$RefreshButton.Text = "Refresh Data"
	$RefreshButton.Add_Click({
		RefreshMainFormData        
	})
	$RefreshButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($RefreshButton)
 
	$CreateSegmentButton = New-Object System.Windows.Forms.Button
	$CreateSegmentButton.Location = New-Object System.Drawing.Size(((10 + (2 * $MainFormButtonWidth) + (2 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$CreateSegmentButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$CreateSegmentButton.Text = "Create Segment"
	$CreateSegmentButton.Add_Click({
		RefreshMainFormData
		EditSegment
	})
	$CreateSegmentButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($CreateSegmentButton)

	$DeleteSegmentButton = New-Object System.Windows.Forms.Button
	$DeleteSegmentButton.Location = New-Object System.Drawing.Size(((10 + (3 * $MainFormButtonWidth) + (3 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$DeleteSegmentButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$DeleteSegmentButton.Text = "Delete Segment"
	$DeleteSegmentButton.Add_Click({
		$ConnectTenantButton.Enabled = $false
		$RefreshButton.Enabled = $false
		$CreateSegmentButton.Enabled = $false
		$DeleteSegmentButton.Enabled=$false
		$CreatePolicyButton.Enabled = $false
		$DeletePolicyButton.Enabled = $false
		$StartApplicationButton.Enabled = $false
		$DeleteSegmentButton.Text = "Deleting..."
		$SelectedSegment= $OrgSegmentDatagrid.Rows[$OrgSegmentDatagrid.CurrentRow.Index].Cells[0].value
		$DeleteResult = [Microsoft.VisualBasic.Interaction]::MsgBox("Delete Segment $($SelectedSegment)?",'YesNoCancel,Question', "Delete Segment?")
		if ($DeleteResult -eq "Yes") 
		{ 
			Get-OrganizationSegment |Where {$_.Name -like $SelectedSegment}|Remove-OrganizationSegment -confirm:$false
			RefreshMainFormData 
		}
		$ConnectTenantButton.Enabled = $true
		$RefreshButton.Enabled = $true
		$CreateSegmentButton.Enabled = $true
		$DeleteSegmentButton.Enabled=$true
		$CreatePolicyButton.Enabled = $true
		$DeletePolicyButton.Enabled = $true
		$StartApplicationButton.Enabled = $true
		$DeleteSegmentButton.Text = "Delete Segment"
	})
	$DeleteSegmentButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($DeleteSegmentButton)

	$CreatePolicyButton = New-Object System.Windows.Forms.Button
	$CreatePolicyButton.Location = New-Object System.Drawing.Size(((10 + (4 * $MainFormButtonWidth) + (4 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$CreatePolicyButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$CreatePolicyButton.Text = "Create Policy"
	$CreatePolicyButton.Add_Click({
		RefreshMainFormData
		EditPolicy
	})
	$CreatePolicyButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($CreatePolicyButton)

	$DeletePolicyButton = New-Object System.Windows.Forms.Button
	$DeletePolicyButton.Location = New-Object System.Drawing.Size(((10 + (5 * $MainFormButtonWidth) + (5 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$DeletePolicyButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$DeletePolicyButton.Text = "Delete Policy"
	$DeletePolicyButton.Add_Click({
		$ConnectTenantButton.Enabled = $false
		$RefreshButton.Enabled = $false
		$CreateSegmentButton.Enabled = $false
		$DeleteSegmentButton.Enabled=$false
		$CreatePolicyButton.Enabled = $false
		$DeletePolicyButton.Enabled = $false
		$StartApplicationButton.Enabled = $false
		$DeletePolicyButton.Text = "Deleting..."
		$SelectedPolicy= $IBPolicyDatagrid.Rows[$IBPolicyDatagrid.CurrentRow.Index].Cells[0].value
		$DeleteResult = [Microsoft.VisualBasic.Interaction]::MsgBox("Delete Policy $($SelectedPolicy)?",'YesNoCancel,Question', "Delete Policy?")
		if ($DeleteResult -eq "Yes") 
		{ 
			$OldPolicy = Get-InformationBarrierPolicy | Where {$_.Name -like $SelectedPolicy}
			Set-InformationBarrierPolicy -identity $OldPolicy.GUID -state inactive
			Remove-InformationBarrierPolicy -identity $OldPolicy.GUID -confirm:$false
			RefreshMainFormData 
		}
		$ConnectTenantButton.Enabled = $true
		$RefreshButton.Enabled = $true
		$CreateSegmentButton.Enabled = $true
		$DeleteSegmentButton.Enabled=$true
		$CreatePolicyButton.Enabled = $true
		$DeletePolicyButton.Enabled = $true
		$StartApplicationButton.Enabled = $true
		$DeletePolicyButton.Text = "Delete Policy"
	})
	$DeletePolicyButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($DeletePolicyButton)

	$StartApplicationButton = New-Object System.Windows.Forms.Button
	$StartApplicationButton.Location = New-Object System.Drawing.Size(((10 + (6 * $MainFormButtonWidth) + (6 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$StartApplicationButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$StartApplicationButton.Text = "Start Policy Application"
	$StartApplicationButton.Add_Click({
		$errorcount=$error.count
		Start-InformationBarrierPoliciesApplication
		if ($error.count -gt $errorcount) 
		{
			[Microsoft.VisualBasic.Interaction]::MsgBox("Error: $(($error[0] -split "Status:")[0])" ,'Exclamation', "Error.")
		}
	})
	$StartApplicationButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($StartApplicationButton)

	$HelpButton = New-Object System.Windows.Forms.Button
	$HelpButton.Location = New-Object System.Drawing.Size(((10 + (7 * $MainFormButtonWidth) + (7 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$HelpButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$HelpButton.Text = "Help"
	$HelpButton.Add_Click({
		ShowHelp
	})
	$HelpButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($HelpButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(((10 + (8 * $MainFormButtonWidth) + (8 * $MainFormButtonSpacing)) + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 110))
	$CancelButton.Size = New-Object System.Drawing.Size($MainFormButtonWidth,$MainFormButtonHeight)
	$CancelButton.Text = "Quit"
	$CancelButton.Add_Click({
		$mainForm.Close()
	})
	$CancelButton.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($CancelButton)

	#TeamsAdmin LinkLabel
	$TeamsAdminLinkLabel = New-Object System.Windows.Forms.LinkLabel
	$TeamsAdminLinkLabel.Location = New-Object System.Drawing.Size(10,($mainForm.height - 60)) 
	$TeamsAdminLinkLabel.Size = New-Object System.Drawing.Size(200,20)
	$TeamsAdminLinkLabel.text = "http://www.TeamsAdmin.com"
	$TeamsAdminLinkLabel.add_Click({Start-Process $TeamsAdminLinkLabel.text})
	$TeamsAdminLinkLabel.Anchor = 'Bottom, Left'
	$mainForm.Controls.Add($TeamsAdminLinkLabel)
   
	[void] $mainForm.ShowDialog()

}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$global:connected=$false
MainForm
