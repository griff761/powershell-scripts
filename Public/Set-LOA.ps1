# Script to process LOA on a user account
# Written by Griffin Rodgers

#Requires -RunAsAdministrator

function Set-LOA {
    # Check if the user already has the AzureAD Powershell module
    InstallAzureAD

    # Collect information needed to process the LOA from the script user
    do {
        $Username = Read-Host "Please enter the username of the user on LOA"
        $Requestor = Read-Host "Please enter the name of the LOA requestor"
        $Initials = Read-Host "Please enter your initials"
        $Ticket = Read-Host "Please enter the ticket number"
        $Password = Read-Host "Please enter a randomly generated password for the disabled user account(must still meet password policy)" -AsSecureString
        $User = Get-ADUser -Identity $Username
        $Today = (Get-Date).ToString('MM/dd/yy')
        $LOA = "LOA as per $Requestor on $Today $Ticket $Initials"
        $OOO = Read-Host "Please enter the Out of Office message (press enter if there is no OOO message)"

        # Have the user confirm the data is correct
        $DataConfirmation = Read-Host "Please review the following data: `r`n $User `r`n Requestor = $Requestor `r`n Initials = $Initials `r`n Ticket = $Ticket `r`n Date = $Today `r`n OOO = $OOO `r`n Is this data correct? Enter 'y' or 'n'"

        while ("y", "Y", "n", "N" -notcontains $DataConfirmation ) {
            $DataConfirmation = Read-Host "Please review the following data: `r`n $User `r`n Requestor = $Requestor `r`n Initials = $Initials `r`n Ticket = $Ticket `r`n Date = $Today `r`n OOO = $OOO `r`n Is this data correct? Enter 'y' or 'n'"
            # Will loop the command again if the user enters an invalid value

        }

    }while ("y", "Y" -notcontains $DataConfirmation)

    # Disables the account, sets a random password, sets account to expire, moves the user to Disabled-Users OU, sets LOA message in description.
    Set-ADAccountPassword $User -Reset -NewPassword $Password
    Set-ADAccountExpiration $User -DateTime ((Get-Date).AddDays(-1).ToString('MM/dd/yyyy'))
    Disable-ADAccount $User
    Move-ADObject $User -TargetPath "OU=Users_Disabled_Exceptions,OU=Users-Disabled,OU=Users" # Set proper target path for organization
    Get-ADUser -Identity $Username -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$LOA $($_.Description)" }

    # Sets OOO message if one is required
    if ([string]::IsNullOrWhitespace($OOO)) {
        Write-Host "No Out of Office message was set"
    }
    else {

        InstallExchangeOnline
    
        # Connect to Exchange Online and set the OOO message
        CheckExchangeOnlineRunning
        Set-CloudMailboxAutoReplyConfiguration $username –AutoReplyState Enabled –ExternalMessage $OOO –InternalMessage $OOO
    }

    # Connect to Azure AD remote powershell
    CheckAzureADRunning

    # Find user in Azure AD and revoke all sessions
    $ObjectID = $(Get-AzureADUser -SearchString "$User@contoso.com").ObjectId

    if ($null -eq $ObjectID) {
        $ObjectID = $(Get-AzureADUser -SearchString "$User").ObjectId
    }

    Revoke-AzureADUserAllRefreshToken -ObjectId $ObjectID

    Write-Host "$Username is now set for LOA. Press enter to exit the script" -ForegroundColor Green
    Read-Host
}
