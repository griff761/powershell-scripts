function Enable-DLForO365Group
{

    # Script to enable an existing O365/Teams Group for DL functionality and subscribe all members
    # Written by Griffin Rodgers

    #Requires -RunAsAdministrator

    InstallExchangeOnline

    do {
        
        $Group = Read-Host -Prompt "Please enter the name of the O365/Teams Group you want to enable DL funcionality for"

        # Have the user confirm the data is correct
        $DataConfirmation = Read-Host "The group name you have entered is '$Group' `r`n Is this correct? Enter 'y' or 'n'"

        while("y","Y","n", "N" -notcontains $DataConfirmation )
        {
        $DataConfirmation = Read-Host "The group name you have entered is $Group `r`n Is this correct? Enter 'y' or 'n'"
        # Will loop the command again if the user enters an invalid value

        }

    }while ("y","Y" -notcontains $DataConfirmation)

    CheckExchangeOnlineRunning # Connect to remote Exchange Online powershell session

    $Members = Get-CloudUnifiedGroupLinks -Identity $Group -LinkType Member # Get all the current members of the group

    Add-CloudUnifiedGroupLinks -Identity $Group -LinkType subscriber -Links $Members.Alias # Subscribe all current members
    Set-CloudUnifiedGroup -Identity $Group -HiddenFromAddressListsEnabled $false # Make the group visible in the address book

    Read-Host "If there are no errors, you have successfully enabled DL functionality for this team! Please press enter to close the script"
}
