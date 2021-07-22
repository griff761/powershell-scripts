# New User Mailbox creation script created by Griffin Rodgers
#Requires -RunAsAdministrator

function New-UserAccount {

    # All necessary values to be collected by the user to create the Shared Mailbox. Will loop the script again if the user reviews the data and notices an error.
    do {

        # Ask user if the new user is BCM or Non-BCM
        $confirmation = Read-Host "Is the user a full time employee or Contractor, Temp, etc? Type 'y' (without quotes) for FTE or type 'n'(without quotes) for Non-FTE"

        while ("y", "Y", "n", "N" -notcontains $confirmation ) {
            $confirmation = Read-Host "Is the user BCM(full time employee) or Non-BCM(Contractor, Temp, etc)? Type 'y' (without quotes) for BCM or type 'n'(without quotes) for Non-BCM"
            # Will loop the command again if the user enters an invalid value
        }

        # If yes, set the Organizational Unit to FTE
        if (($confirmation -eq 'y') -or ($confirmation -eq 'Y')) {
            $OrganizationalUnit = "" #OU will need to be specified
        }

        # If no, set the Organizational Unit to Non-FTE
        else {
            $OrganizationalUnit = "" #OU will need to be specified
        }


        # Collect new user infromation
        $FirstName = Read-Host -Prompt "Enter the user's first name"
        $Initial = Read-Host -Prompt "Enter the user's middle inital (if not applicable, just hit enter)"
        $LastName = Read-Host -Prompt "Enter the user's last name"
        $MirrorID = Read-Host -Prompt "Enter the username of the mirror user"
        $Credentials = Get-Credential -Message "Please enter the username and password of the new user"
        $TicketNumber = Read-Host -Prompt "Enter the ticket number for the account creation"
        $AccountRequestor = Read-Host -Prompt "Enter the name of the person who put in the request"
        $Initials = Read-Host -Prompt "Enter the HD Engineer's initials"
        $DisplayName = "$FirstName $LastName"
        $newuser = $Credentials.UserName
        $UPN = $Credentials.UserName
        $Today = (Get-Date).ToString('MM/dd/yy')
        $Description = "As Per $AccountRequestor on $Today Req $TicketNumber $Initials"

        # Have the user confirm the data is correct
        $DataConfirmation = Read-Host "Please review the following data: `r`n First Name = $FirstName `r`n Initial = $Initial `r`n Last Name = $LastName `r`n Display Name = $DisplayName `r`n Organization Unit = $OrganizationalUnit `r`n Mirror Username = $MirrorID `r`n Ticket # = $TicketNumber `r`n Account Requestor = $AccountRequestor `r`n Engineer Initials = $Initials `r`n Is the above data correct (enter 'y' or 'n' without quotes): "

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Please review the following data: `r`n First Name = $FirstName `r`n Initial = $Initial `r`n Last Name = $LastName `r`n Display Name = $DisplayName `r`n Organization Unit = $OrganizationalUnit `r`n Mirror Username = $MirrorID `r`n Ticket # = $TicketNumber `r`n Account Requestor = $AccountRequestor `r`n Engineer Initials = $Initials `r`n Is the above data correct (enter 'y' or 'n' without quotes): "
            # Will loop the command again if the user enters an invalid value

        }

    }while ("y", "Y" -notcontains $DataConfirmation)

    # Script will only get to this point once the user confirms that all of the data they entered is correct

    CheckExchangeRemotelyRunning # Connect Powershell to Exchange on-prem to receive necessary commands

    New-RemoteMailbox -Name $DisplayName -OnPremisesOrganizationalUnit $OrganizationalUnit -UserPrincipalName "$UPN@berkadia.com" -FirstName $FirstName -LastName $LastName -Initials $Initial -Password $Credentials.Password -ResetPasswordOnNextLogon $false -Archive
    # The above command creates a mailbox from an existing AD account

    # Begin group mirroring process

    $outputfile = "" #Output path where group membership is to be saved will need to be put here.

    $output = @()

    $groups = (Get-ADUser -Identity $MirrorID -Properties memberOf | Select-Object memberOf).memberOf
    foreach ($group in $groups) {
	    $groupinfo = Get-ADGroup -Identity $group -Properties info,managedBy,Description | Select-Object Name,GroupCategory,@{Name='ManagedBy';Expression={(Get-ADUser $_.managedBy).GivenName,(Get-ADUser $_.managedBy).Surname}},Info
    if ($groupinfo.info.length -gt 0) {
		    $groupinfo.info = $groupinfo.info.replace("`n","|")
	    } else {
		    $groupinfo.info = " "
	    }
	    #write-output $groupinfo
	    #$groupinfo | Export-CSV -Append -NoTypeInformation -Path $outputfile 
	    $output += $groupinfo
    }

    $output | Sort-Object -Property Name | Export-CSV -NoTypeInformation -Path $outputfile

    #Install AzureAD module if not installed
    InstallAzureAD

    CheckAzureADRunning

    #Get the username for the user account
    $UserAAD = $MirrorID + "@berkadia.com"

    #Get the Azure account ObjectID
    $AADUser = $(Get-AzureADUser -Filter "UserPrincipalName eq '$UserAAD'").ObjectId

    $Results = Get-AzureADUserMembership -ObjectId $AADUser #| Select-Object -Property  DisplayName,ObjectType,Description,ObjectId,GroupTypes | Export-Csv -Path C:\Temp\$User.csv

    $GroupList = @()

    ForEach ($result in $Results){

    if ($result.ObjectType -eq "Role")
    {
        continue
    }
    $Group = Get-AzureADMSGroup -Id $result.ObjectID 
    $Manager = Get-AzureADGroupOwner -ObjectId $result.ObjectId
    Add-Member -InputObject $Group -Name "ManagedBy" -Value $Manager.DisplayName -MemberType NoteProperty
    Add-Member -InputObject $Group -Name "Info" -Value "Azure AD Group" -MemberType NoteProperty
    $GroupList += $Group | Select-Object @{name="Name";expression={$_.DisplayName}},@{name="GroupCategory";expression={$_.GroupTypes -join ", "}},SecurityEnabled,@{name="ManagedBy";expression={$_.ManagedBy -join ", "}}, OnPremisesSyncEnabled, Info | Where-Object {($_.GroupCategory -notcontains "Unified") -and ($_.GroupCategory -notcontains "DynamicMembership") -and ($_.OnPremisesSyncEnabled -notcontains "True")}
    }

    ForEach ($GroupItem in $GroupList){

        if ($GroupItem.SecurityEnabled -eq "True"){
        $GroupItem.GroupCategory = "Security"
        }

        else{
        $GroupItem.GroupCategory = "Distribution"
        }

        $GroupItem.PSObject.Properties.Remove('OnPremisesSyncEnabled')
        $GroupItem.PSObject.Properties.Remove('SecurityEnabled')
    
    }

    $GroupList | Export-Csv -Path $outputfile -Append -NoTypeInformation

    # Clear variable
    $UserAD = $null

    # Wait for new user account to replicate to AD

        Write-Host "This part of the script may take 10+ minutes to run. The user account needs to replicate to AD. The script will check if the account has replicated every 60 seconds." -ForegroundColor Red
        do {
            Start-Sleep -Seconds 60

            try {
                $UserAD = Get-ADUser $UPN
            }

            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $UserAD = $null
            }

            if ($UserAD -eq $null)
            {
                Write-Host "New user account has not replicated to AD yet. Will try again in 60 seconds." -ForegroundColor Red
            }
        }
        while ($UserAD -eq $null)

        $UserAD | Set-ADUser -Description $Description

        foreach ($group in $groups)
        {
            Add-ADGroupMember -Identity $group -Members $newuser
        }

        Write-Host "User and mailbox have been created! Please continue with the next steps outlined in the documentation."
}
