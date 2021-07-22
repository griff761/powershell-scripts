#This script will hide a user's email address in the GAL
#Created by Griffin Rodgers

#Requires -RunAsAdministrator

function Hide-EmailGAL {

    CheckExchangeRemotelyRunning

    $User = Read-Host -Prompt 'Enter the email address for the account'

    Set-RemoteMailbox $User -HiddenFromAddressListsEnabled $true

    Read-Host -Prompt 'Press Enter to exit'
}
