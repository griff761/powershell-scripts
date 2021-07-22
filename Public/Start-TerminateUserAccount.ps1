# Powershell Script to terminate a user
# Authored by Griffin Rodgers

#Requires -RunAsAdministrator

function Start-TerminateUserAccount {

    Function TermInputForm {

        function CheckAllBoxes {
            if ( ($textbox1.Text.Length -and $textbox2.Text.Length -and $textBox3.Text.Length -and $textBox4.Text.Length -and $textBox5.Text.Length) -gt 0) {
                $okButton.Enabled = $true
            }
            else {
                $okButton.Enabled = $false
            }
        }
        
        
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Termination Info'
        $form.Size = New-Object System.Drawing.Size(300, 475)
        $form.StartPosition = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(75, 395)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Text = 'OK'
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $okButton
        $form.Controls.Add($okButton)
        $okButton.Enabled = $false
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(150, 395)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $cancelButton.Text = 'Cancel'
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $cancelButton
        $form.Controls.Add($cancelButton)
        
        $label1 = New-Object System.Windows.Forms.Label
        $label1.Location = New-Object System.Drawing.Point(10, 20)
        $label1.Size = New-Object System.Drawing.Size(280, 30)
        $label1.Text = 'Please enter the username of the user being terminated:'
        $form.Controls.Add($label1)
        
        $textBox1 = New-Object System.Windows.Forms.TextBox
        $textBox1.Location = New-Object System.Drawing.Point(10, 50)
        $textBox1.Size = New-Object System.Drawing.Size(260, 20)
        $textBox1.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox1)
        
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(10, 75)
        $label2.Size = New-Object System.Drawing.Size(280, 20)
        $label2.Text = 'Please enter the name of the termination requestor:'
        $form.Controls.Add($label2)
        
        $textBox2 = New-Object System.Windows.Forms.TextBox
        $textBox2.Location = New-Object System.Drawing.Point(10, 95)
        $textBox2.Size = New-Object System.Drawing.Size(260, 20)
        $textBox2.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox2)
        
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10, 120)
        $label3.Size = New-Object System.Drawing.Size(280, 20)
        $label3.Text = 'Please enter your initials:'
        $form.Controls.Add($label3)
        
        $textBox3 = New-Object System.Windows.Forms.TextBox
        $textBox3.Location = New-Object System.Drawing.Point(10, 140)
        $textBox3.Size = New-Object System.Drawing.Size(260, 20)
        $textBox3.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox3)
        
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10, 165)
        $label4.Size = New-Object System.Drawing.Size(280, 20)
        $label4.Text = 'Please enter the ticket number:'
        $form.Controls.Add($label4)
        
        $textBox4 = New-Object System.Windows.Forms.TextBox
        $textBox4.Location = New-Object System.Drawing.Point(10, 185)
        $textBox4.Size = New-Object System.Drawing.Size(260, 20)
        $textBox4.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox4)
    
        $label5 = New-Object System.Windows.Forms.Label
        $label5.Location = New-Object System.Drawing.Point(10, 210)
        $label5.Size = New-Object System.Drawing.Size(280, 40)
        $label5.Text = 'Please enter a randomly generated password for the disabled user account(must still meet password policy):'
        $form.Controls.Add($label5)
        
        $textBox5 = New-Object System.Windows.Forms.MaskedTextBox
        $textBox5.PasswordChar = '*'
        $textBox5.Location = New-Object System.Drawing.Point(10, 255)
        $textBox5.Size = New-Object System.Drawing.Size(260, 20)
        $textBox5.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox5)
    
        $label6 = New-Object System.Windows.Forms.Label
        $label6.Location = New-Object System.Drawing.Point(10, 280)
        $label6.Size = New-Object System.Drawing.Size(280, 30)
        $label6.Text = 'Please enter the Out of Office message (leave blank if there is no OOO message):'
        $form.Controls.Add($label6)
        
        $textBox6 = New-Object System.Windows.Forms.TextBox
        $textBox6.Location = New-Object System.Drawing.Point(10, 310)
        $textBox6.Size = New-Object System.Drawing.Size(260, 20)
        $textBox6.add_TextChanged( { CheckAllBoxes })
        $form.Controls.Add($textBox6)
        
        $form.Topmost = $true
        
        $form.Add_Shown( { $textBox1.Select() })
        $result = $form.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $x = @()
            $x += $textBox1.Text
            $x += $textBox2.Text
            $x += $textBox3.Text
            $x += $textBox4.Text
            $x += $textBox5.Text | ConvertTo-SecureString -AsPlainText -Force
            $x += $textBox6.Text
            $x
        }
        
        else { $x = $null }
        
    }

    InstallAzureAD

    InstallExchangeOnline

    InstallImportExcel

    # Establish necessary remote powershell sessions
    CheckExchangeRemotelyRunning

    # Collects info relevant to the termination
    do {

        $Info = TermInputForm

        if ($null -eq $Info) {
            throw "User cancelled data entry! Exiting script."
        }

        $Username = $Info[0]
        $Requestor = $Info[1]
        $Initials = $Info[2]
        $Ticket = $Info[3]
        $Password = $Info[4]
        $OOO = $Info[5]

        # $DisplayName = Read-Host "Please enter the display name of the user in AD"
        # $Username = Read-Host "Please enter the username of the user being terminated"
        # $Requestor = Read-Host "Please enter the name of the termination requestor"
        # $Initials = Read-Host "Please enter your initials"
        # $Ticket = Read-Host "Please enter the ticket number (BITXXXXXX)"
        # $Password = Read-Host "Please enter a randomly generated password for the disabled user account(must still meet Berkadia password policy)" -AsSecureString
        $User = Get-ADUser -Identity $Username -Properties cn, whenCreated, description
        $Today = (Get-Date).ToString('MM/dd/yy')
        $Disabled = "Disabled as per $Requestor on $Today $Ticket $Initials"
        # $OOO = Read-Host "Please enter the Out of Office message (press enter if there is no OOO message)"

        # Have the user confirm the data is correct
        $DataConfirmation = Read-Host "Please review the following data: `r`n $User `r`n Requestor = $Requestor `r`n Initials = $Initials `r`n Ticket = $Ticket `r`n Date = $Today `r`n OOO = $OOO `r`n Is this data correct? Enter 'y' or 'n'"

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Please review the following data: `r`n $User `r`n Requestor = $Requestor `r`n Initials = $Initials `r`n Ticket = $Ticket `r`n Date = $Today `r`n OOO = $OOO `r`n Is this data correct? Enter 'y' or 'n'"
            # Will loop the command again if the user enters an invalid value

        }

    }while ("y", "Y" -notcontains $DataConfirmation)

    # Disables the account, sets a random password, sets account to expire, moves the user to Disabled-Users OU, sets Disabled message in description, hides mailbox.
    Set-ADAccountPassword $User -Reset -NewPassword $Password
    Set-ADAccountExpiration $User -DateTime ((Get-Date).AddDays(-1).ToString('MM/dd/yyyy'))
    Disable-ADAccount $User
    Move-ADObject $User -TargetPath "OU=Users-Disabled,OU=Users" #Must set proper Target Path for organization
    Get-ADUser -Identity $Username -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$Disabled $($_.Description)" }
    Set-RemoteMailbox $Username -HiddenFromAddressListsEnabled $true

    # Sets OOO message if one is reqauired
    if ([string]::IsNullOrWhitespace($OOO)) {
        Write-Host "No Out of Office message was set"
    }

    else {
        # Check if Exchange Online is running, connect if not.
        CheckExchangeOnlineRunning

        Set-CloudMailboxAutoReplyConfiguration $username –AutoReplyState Enabled –ExternalMessage $OOO –InternalMessage $OOO
    }

    # Runs script to gather user groups
    Start-Sleep -Seconds 30
    $User = Get-ADUser -Identity $Username -Properties cn, whenCreated, description
    SaveTerminatedUserGroups -ADUser $User

    do {
        $DataConfirmation = Read-Host "Have you checked the excel spreadsheet at 'INSERT CORRECT PATH HERE' to make sure the groups are saved? Enter 'y' or 'n'"

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Have you checked the excel spreadsheet at 'INSERT CORRECT PATH HERE' to make sure the groups are saved? Enter 'y' or 'n'"
            # Will loop the command again if the user enters an invalid value

        }
    }while ("y", "Y" -notcontains $DataConfirmation)

    # Remove the user from all groups they are a member of
    Get-AdPrincipalGroupMembership -Identity $Username | Where-Object -Property Name -Ne -Value 'Domain Users' | Remove-AdGroupMember -Members $User

    # Check if AzureAD module is running, connect if not
    CheckAzureADRunning
    
    # Connect to AzureAD remote powershell session
    # Connect-AzureAD

    $ObjectID = $(Get-AzureADUser -SearchString "$Username@contoso.com").ObjectId

    if ($null -eq $ObjectID) {
        $ObjectID = $(Get-AzureADUser -SearchString "$Username").ObjectId
    }

    try {
        Remove-AzureADGroupMember -ObjectId 0e3717a5-51e2-4544-a8fd-90212f61aa31 -MemberId $ObjectID
        Remove-AzureADGroupMember -ObjectId 16219c47-bf89-4508-8f81-d5b70332d969 -MemberId $ObjectID
    }
    catch [Microsoft.Open.AzureAD16.Client.ApiException] {
        Write-Host -ForegroundColor Green "User not a member of Intune and/or Intune WiFi groups"
    }

    Revoke-AzureADUserAllRefreshToken -ObjectId $ObjectID

    # Lists all groups owned by the terminated user

    Write-Host "$Username is now disabled. Below are the objects in AD owned by the terminated user. If any objects are listed below, contact the terminated user's manager to find out who should be the new owner"

    $DisplayName = $User.CN

    $results = @(Get-ADObject -LDAPFilter "(|(managedBy=*)(manager=*))" -Properties manager, managedBy | `
            Where-Object { ($_.managedBy -like "*$DisplayName*" -or $_.info -like "*$DisplayName*" -or $_.manager -like "*$DisplayName*") -and ($_.DistinguishedName -notlike ("CN=*,OU=*Users*,OU=Users")) } | `
            Select-Object name, objectclass, @{label = 'Manager'; expression = { $_.manager -replace '^CN=|,.*$' } }, @{label = 'ManagedBy'; expression = { $_.managedBy -replace '^CN=|,.*$' } } | `
            Sort-Object -Property ManagedBy, Manager, objectClass, name) #Correct path will be required

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
