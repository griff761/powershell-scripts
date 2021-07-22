# Script to add senders and/or domains to the Safe Senders list of a mail account (User accounts, shared mailboxes)
# Created by Griffin Rodgers

#Requires -RunAsAdministrator

Function Add-SafeSendersAndDomains {

InstallExchangeOnline

CheckExchangeOnlineRunning

do {
    Write-Host "Please select the CSV file containing the list of domains/email addresses you wish to add to the mailbox's safe sender list" -ForegroundColor Green
    $Path = Get-FileName -Title "Please select the CSV file containing the list of domains/email addresses you wish to add to the mailbox's safe sender list"
    $Email = Read-Host -Prompt "Please enter the email address of the account which you wish to add entries to the safe sender list"
    $EmailEXO = Get-CloudMailbox $Email

    # Have the user confirm the data is correct
    $DataConfirmation = Read-Host "Please review the following data: `r`n Path = $Path `r`n Account = $EmailEXO `r`n Is this data correct? Enter 'y' or 'n'"

    while("y","Y","n", "N" -notcontains $DataConfirmation )
    {
    $DataConfirmation = Read-Host "Please review the following data: `r`n Path = $Path `r`n Account = $EmailEXO `r`n Is this data correct? Enter 'y' or 'n'"
    # Will loop the command again if the user enters an invalid value

    }
}while ("y","Y" -notcontains $DataConfirmation)

$Domains = Import-CSV -Path $Path -Header Domains

ForEach($Domain in $Domains)
{
    Set-CloudMailboxJunkEmailConfiguration -Identity $EmailEXO.UserPrincipalName -TrustedSendersAndDomains @{Add=$Domain.Domains} 
}

Write-Host "As long as there are no errors, the senders/domains are now added to the safe sender list of the account."

}