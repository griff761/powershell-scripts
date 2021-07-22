#Shared Mailbox creation script created by Griffin Rodgers
#Requires -RunAsAdministrator

function New-SharedMailbox {

    InstallExchangeOnline

    # All necessary values to be collected by the user to create the Shared Mailbox
    do {
        $ADAccount = Read-Host -Prompt "Enter the username for the disabled mailbox account in AD (Ex: For 'Script Test', the username would be 'stest')"
        $DisplayName = Read-Host -Prompt 'Enter the display name of the Shared Mailbox'
        $Alias = Read-Host -Prompt 'Enter the alias of the Shared Mailbox (Display name without spaces)'
        $SecurityGroup = "MBX $DisplayName - Full"
        $Ticket = Read-Host -Prompt "Enter the ticket number for this request"
        $Description = "Full and Send As permissions to $DisplayName Mailbox"
        $OwnerUserName = Read-Host -Prompt "Please enter the username of the owner of the mailbox"
        $Requestor = Read-Host -Prompt "Please enter the name of the requestor"
        $Today = (Get-Date).ToString('MM/dd/yy')
        $Initials = Read-Host -Prompt "Please enter the HD Engineer's initials"
        $Owner = Get-ADUser $OwnerUserName
        $OwnerName = $Owner.Name

        # Have the user confirm the data is correct
        $DataConfirmation = Read-Host "Please review the following data: `r`n Username = $ADAccount `r`n Display Name = $DisplayName `r`n Alias = $Alias `r`n Security Group = $SecurityGroup `r`n Ticket Number = $Ticket `r`n Owner = $OwnerName `r`n Requestor = $Requestor `r`n Date = $Today `r`n Initials = $Initials `r`n Is this data correct? Enter 'y' or 'n'"

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Please review the following data: `r`n Username = $ADAccount `r`n Display Name = $DisplayName `r`n Alias = $Alias `r`n Security Group = $SecurityGroup `r`n Ticket Number = $Ticket `r`n Owner = $OwnerName `r`n Requestor = $Requestor `r`n Date = $Today `r`n Initials = $Initials `r`n Is this data correct? Enter 'y' or 'n'"
            # Will loop the command again if the user enters an invalid value

        }
    }while ("y", "Y" -notcontains $DataConfirmation)


    CheckExchangeRemotelyRunning # Connect Powershell to Exchange on-prem to receive necessary commands

    $securityGroupTrim = ($SecurityGroup -replace '\s','') + "@berkadia.com"

    # Create MBX Security Group to control access
    New-ADGroup -DisplayName $SecurityGroup -Name $SecurityGroup -GroupCategory Security -GroupScope Universal -Path "OU=Groups,OU=Berkadia,DC=gmaccm,DC=com" -OtherAttributes @{info = "Owner: $OwnerName"; mail = $securityGroupTrim} -ManagedBy $Owner -Description $Description

    # Create Mailbox and Disabled Account
    New-RemoteMailbox -DisplayName $Displayname -UserPrincipalName $ADAccount@berkadia.com -Alias $Alias -Name $DisplayName -FirstName $DisplayName -Archive -OnPremisesOrganizationalUnit "gmaccm.com/Berkadia/Exchange/Shared Mailboxes" -AccountDisabled

    # Pause for a moment to allow time for the mailbox to be found
    Start-Sleep -Seconds 60
    $Mailbox = Get-RemoteMailbox "$Alias@berkadia.com"

    # These two lines of code do the job of the confirm-EA5 script
    $Base64 = [System.Convert]::ToBase64String($Mailbox.Guid.ToByteArray())
    $Mailbox | Set-RemoteMailbox -CustomAttribute14 "X" -CustomAttribute5 $Base64

    # Clear out these variables
    $DisabledUser = $null
    $ADSecurityGroup = $null

    # Try to find the user account and security gorup in AD every 60 seconds. Loops until the account and security group is found.
    Write-Host "This part of the script may take 10+ minutes to run. The disabled user account needs to replicate to AD. The script will check if the account has replicated every 60 seconds." -ForegroundColor Red
    do {
        Start-Sleep -Seconds 60

        try {
            $DisabledUser = Get-ADUser $ADAccount
        }

        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $DisabledUser = $null
        }

        $ADSecurityGroup = Get-ADGroup $SecurityGroup -ErrorAction SilentlyContinue

        if (($DisabledUser -eq $null) -or ($ADSecurityGroup -eq $null))
        {
            Write-Host "Disabled user account has not replicated to AD yet. Will try again in 60 seconds." -ForegroundColor Red
        }
    }
    while (($DisabledUser -eq $null) -or ($ADSecurityGroup -eq $null))

    # Set manager of the mailbox, fill out Notes tab under Telephone, and fill out the descripiton. Mail-enable the MBX security group
    $DisabledUser | Set-ADUser -Replace @{Info = "Owner: $OwnerName" } -Manager $Owner -Description "Disabled Mailbox Acct, Req $Requestor $Today $Ticket $Initials"
    Enable-DistributionGroup -Identity $SecurityGroup


    # This command gives the MBX Security Group Send As rights to the mailbox in AD.
    Get-RemoteMailbox "$Alias@berkadia.com" | Add-ADPermission -User $SecurityGroup -ExtendedRights "Send As"

    # Connect to Exchange Online to receive necessary commands
    CheckExchangeOnlineRunning

    # Clear out variables
    $CloudMailbox = $null
    $CloudSecurityGroup = $null

    # Search for the mailbox and the MBX Security Group in the cloud every 10 minutes. Loops until both are found.
    Write-Host "This part of the script can take 30+ minutes to complete as ADSync needs to run, which happens every 30 minutes. This script will try to find the mailbox in the cloud every 10 minutes." -ForegroundColor Red
    do {
        Start-Sleep -Seconds 600
        $CloudMailbox = Get-cloudMailbox "$Alias@berkadia.com" -ErrorAction SilentlyContinue
        $CloudSecurityGroup = Get-cloudGroup $SecurityGroup -ErrorAction SilentlyContinue

        if (($CloudMailbox -eq $null) -or ($CloudSecurityGroup -eq $null)) {
            Write-Host "Mailbox has not synced to Exchange Online yet. Will try again in 10 Minutes." -ForegroundColor Red
        }
    }
    while (($CloudMailbox -eq $null) -or ($CloudSecurityGroup -eq $null))

    Write-Host "Mailbox Found! Delegating Full Access to $SecurityGroup" -ForegroundColor Green

    $CloudMailbox | Set-cloudMailbox -Type Shared # Set mailbox as shared mailbox
    $CloudMailbox | Add-cloudMailboxPermission -User $CloudSecurityGroup.Id -AccessRights FullAccess -InheritanceType All -AutoMapping $true # Grant Full Access to the security group
    $CloudMailbox | Add-cloudRecipientPermission -Trustee $CloudSecurityGroup.Id -AccessRights SendAs # Grant Send As permission to the security group

    Write-Host "Mailbox Created! Please press enter to exit the script" -ForegroundColor Green
    Read-Host
}
