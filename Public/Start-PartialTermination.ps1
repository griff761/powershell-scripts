# Written by Griffin Rodgers
# Partial termination script

#Requires -RunAsAdministrator

Function Start-PartialTermination {
    InstallAzureAD
    InstallImportExcel

    # Function that will bring up radio button for the access type
    function Radio_Form {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
        # Set the size of your form
        $Form = New-Object System.Windows.Forms.Form
        $Form.width = 500
        $Form.height = 300
        $Form.Text = "Partial Term Access"
 
        # Set the font of the text to be used within the form
        $Font = New-Object System.Drawing.Font("Segoe", 11)
        $Form.Font = $Font
 
        # Create a group that will contain your radio buttons
        $MyGroupBox = New-Object System.Windows.Forms.GroupBox
        $MyGroupBox.Location = '40,30'
        $MyGroupBox.size = '400,140'
        $MyGroupBox.text = "Which access is required for this partial termination?"
    
        # Create the collection of radio buttons
        $RadioButton1 = New-Object System.Windows.Forms.RadioButton
        $RadioButton1.Location = '20,40'
        $RadioButton1.size = '350,20'
        $RadioButton1.Checked = $false 
        $RadioButton1.Text = "Email only"
        $RadioButton1.Add_Click( {
                $OKButton.Enabled = $true
                $RadioButton3.Checked = $false })
 
        $RadioButton3 = New-Object System.Windows.Forms.RadioButton
        $RadioButton3.Location = '20,70'
        $RadioButton3.size = '350,20'
        $RadioButton3.Checked = $false
        $RadioButton3.Text = "Exchange and Teams"
        $RadioButton3.Add_Click( {
                $OKButton.Enabled = $true
                $RadioButton1.Checked = $false })
 
        # Add an OK button
        # Thanks to J.Vierra for simplifing the use of buttons in forms
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '130,200'
        $OKButton.Size = '100,40' 
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $OKButton.Enabled = $false
 
        #Add a cancel button
        $CancelButton = new-object System.Windows.Forms.Button
        $CancelButton.Location = '255,200'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = "Cancel"
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
 
        # Add all the GroupBox controls on one line
        $MyGroupBox.Controls.AddRange(@($Radiobutton1, $RadioButton3))
 
        # Add all the Form controls on one line 
        $form.Controls.AddRange(@($MyGroupBox, $OKButton, $CancelButton))
 
 
    
        # Assign the Accept and Cancel options in the form to the corresponding buttons
        $form.AcceptButton = $OKButton
        $form.CancelButton = $CancelButton
 
        # Activate the form
        $form.Add_Shown( { $form.Activate() })    
    
        # Get the results from the button click
        $dialogResult = $form.ShowDialog()

        if ($dialogResult -eq "OK") {
            # Check the current state of each radio button and respond accordingly
            if ($RadioButton1.Checked -and (!($RadioButton3.Checked))) {
                $result = "ExchangeOnly"
            }
            elseif ($RadioButton3.Checked = $true) { $result = "ExchangeTeams" }
        }

        else { $result = $null }

        return $result
    }

    # Function that will bring up data entry form for termination
    Function PartialTermInputForm {

        function CheckAllBoxes {
            if ( ($textbox.Text.Length -and $textbox1.Text.Length -and $textbox2.Text.Length -and $textBox3.Text.Length -and $textBox4.Text.Length) -gt 0) {
                $okButton.Enabled = $true
            }
            else {
                $okButton.Enabled = $false
            }
        }
    
    
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    
        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Partial Termination Info'
        $form.Size = New-Object System.Drawing.Size(300, 375)
        $form.StartPosition = 'CenterScreen'
    
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(75, 275)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Text = 'OK'
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $okButton
        $form.Controls.Add($okButton)
        $okButton.Enabled = $false
    
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(150, 275)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $cancelButton.Text = 'Cancel'
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $cancelButton
        $form.Controls.Add($cancelButton)
    
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, 20)
        $label.Size = New-Object System.Drawing.Size(280, 20)
        $label.Text = 'Please enter the display name of the user in AD:'
        $form.Controls.Add($label)
    
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10, 40)
        $textBox.Size = New-Object System.Drawing.Size(260, 20)
        $textBox.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox)
    
        $label1 = New-Object System.Windows.Forms.Label
        $label1.Location = New-Object System.Drawing.Point(10, 65)
        $label1.Size = New-Object System.Drawing.Size(280, 30)
        $label1.Text = 'Please enter the username of the user being terminated:'
        $form.Controls.Add($label1)
    
        $textBox1 = New-Object System.Windows.Forms.TextBox
        $textBox1.Location = New-Object System.Drawing.Point(10, 95)
        $textBox1.Size = New-Object System.Drawing.Size(260, 20)
        $textBox1.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox1)
    
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(10, 120)
        $label2.Size = New-Object System.Drawing.Size(280, 20)
        $label2.Text = 'Please enter the name of the termination requestor:'
        $form.Controls.Add($label2)
    
        $textBox2 = New-Object System.Windows.Forms.TextBox
        $textBox2.Location = New-Object System.Drawing.Point(10, 140)
        $textBox2.Size = New-Object System.Drawing.Size(260, 20)
        $textBox2.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox2)
    
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10, 165)
        $label3.Size = New-Object System.Drawing.Size(280, 20)
        $label3.Text = 'Please enter your initials:'
        $form.Controls.Add($label3)
    
        $textBox3 = New-Object System.Windows.Forms.TextBox
        $textBox3.Location = New-Object System.Drawing.Point(10, 185)
        $textBox3.Size = New-Object System.Drawing.Size(260, 20)
        $textBox3.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox3)
    
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10, 210)
        $label4.Size = New-Object System.Drawing.Size(280, 20)
        $label4.Text = 'Please enter the ticket number:'
        $form.Controls.Add($label4)
    
        $textBox4 = New-Object System.Windows.Forms.TextBox
        $textBox4.Location = New-Object System.Drawing.Point(10, 230)
        $textBox4.Size = New-Object System.Drawing.Size(260, 20)
        $textBox4.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox4)
    
        $form.Topmost = $true
    
        $form.Add_Shown( { $textBox.Select() })
        $result = $form.ShowDialog()
    
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $x = @()
            $x += $textBox.Text
            $x += $textBox1.Text
            $x += $textBox2.Text
            $x += $textBox3.Text
            $x += $textBox4.Text
            $x
        }
    
        else { $x = $null }
    
    }

    # Collects info relevant to the termination
    do {

        $Info = PartialTermInputForm

        if ($null -eq $Info) {
            throw "User cancelled data entry! Exiting script."
        }

        $DisplayName = $Info[0]
        $Username = $Info[1]
        $Requestor = $Info[2]
        $Initials = $Info[3]
        $Ticket = $Info[4]

        $User = Get-ADUser -Identity $Username
        $Today = (Get-Date).ToString('MM/dd/yy')
        $Disabled = "Partial Termination as per $Requestor on $Today $Ticket $Initials"

        # Have the user confirm the data is correct
        $DataConfirmation = Read-Host "Please review the following data: `r`n $User `r`n Requestor = $Requestor `r`n Initials = $Initials `r`n Ticket = $Ticket `r`n Date = $Today `r`n Is this data correct? Enter 'y' or 'n'"

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Please review the following data: `r`n $User `r`n Requestor = $Requestor `r`n Initials = $Initials `r`n Ticket = $Ticket `r`n Date = $Today `r`n Is this data correct? Enter 'y' or 'n'"
            # Will loop the command again if the user enters an invalid value

        }

    }while ("y", "Y" -notcontains $DataConfirmation)

    Get-ADUser -Identity $Username -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$Disabled $($_.Description)" }

    # Runs script to gather user groups
    $User = Get-ADUser -Identity $Username -Properties cn, whenCreated, description
    SaveTerminatedUserGroups -ADUser $User

    # Confirm the user has saved the groups
    do {
        $DataConfirmation = Read-Host "Have you checked the excel spreadsheet at C:\temp to make sure the groups are saved? Enter 'y' or 'n'"

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Have you checked the excel spreadsheet at c:\temp to make sure the groups are saved? Enter 'y' or 'n'"
            # Will loop the command again if the user enters an invalid value

        }
    }while ("y", "Y" -notcontains $DataConfirmation)

    # Remove the user from all groups they are a member of
    Get-AdPrincipalGroupMembership -Identity $Username | Where-Object -Property Name -Ne -Value 'Domain Users' | Remove-AdGroupMember -Members $User

    #Check if user is already connected to AzureAD
    CheckAzureADRunning

    $ObjectID = $(Get-AzureADUser -SearchString "$Username@contoso.com").ObjectId #Will need to change this to match whatever organization.

    if ($null -eq $ObjectID) {
        $ObjectID = $(Get-AzureADUser -SearchString "$Username").ObjectId
    }

    $AccessType = Radio_Form

    if ($null -eq $AccessType) {
        throw "User cancelled! Exiting script."
    }

    elseif ($AccessType -eq "ExchangeOnly") {
        Add-AzureADGroupMember -ObjectId "" -RefObjectId $ObjectID #Specify ObjectID of group that only allows Email access
    }

    elseif ($AccessType -eq "ExchangeTeams") {
        Add-AzureADGroupMember -ObjectId "" -RefObjectId $ObjectID #Specificy ObjectID of group that allows Email and Teams access
    }

    Write-Host "$Username is now partially terminated. Below are the groups owned by the terminated user. If any groups are listed below, contact the terminated user's manager to find out who should be the new owner"

    $DisplayName = $User.CN

    $results = @(Get-ADObject -LDAPFilter "(|(managedBy=*)(manager=*))" -Properties manager, managedBy | `
            Where-Object { ($_.managedBy -like "*$DisplayName*" -or $_.info -like "*$DisplayName*" -or $_.manager -like "*$DisplayName*") -and ($_.DistinguishedName -notlike ("CN=*,OU=*Users*,OU=Users")) } | `
            Select-Object name, objectclass, @{label = 'Manager'; expression = { $_.manager -replace '^CN=|,.*$' } }, @{label = 'ManagedBy'; expression = { $_.managedBy -replace '^CN=|,.*$' } } | `
            Sort-Object -Property ManagedBy, Manager, objectClass, name)

    #write-output $results

    if ($results.Count -gt 0) {
        $usersArray = @()
        $othersArray = @()
        Write-Output "Groups"
        foreach ($result in $results) {
            if (($result.objectclass -eq "group") -and ($result.name -ne "Domain Users")) {
                Write-Output "`t$($result.Name) ($($result.objectClass))" 
            }
            elseif ($result.objectclass -eq "user") {
                $usersArray += $result
            }
            else {
                $othersArray += $result
            }
        }
        Write-Output "`n"

        if ($usersArray.Count -gt 0) {
            Write-Output "Users"
            foreach ($result in $usersArray) {
                Write-Output "`t$($result.Name) ($($result.objectClass))" 
            }
            Write-Output "`n"
        }


        if ($othersArray.Count -gt 0) {
            Write-Output "Other"
            foreach ($result in $othersArray) {
                Write-Output "`t$($result.Name) ($($result.objectClass))" 
            }
            Write-Output "`n"
        }

    }

    Read-Host "Press enter when you are ready to leave the script, make sure you have saved the groups owned by the user before doing so"
}
