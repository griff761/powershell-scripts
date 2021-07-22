#This script will pull all the Azure groups a user belongs to and export it to an Excel spreadsheet
#Created by Griffin Rodgers

#Requires -RunAsAdministrator

Function Pull-ADGroupMembership {

if (Get-Module -ListAvailable -Name "AzureAD") {
        Write-Host 'AzureAD Module already installed'
    }
    else {
        try {
            Install-Module AzureAD 
        }
        catch [Exception] {
            $_.message
            exit
        }
    }

Connect-AzureAD

#Get the user's account information
$User = Read-Host -Prompt 'Enter the username for the account'

#Pipe the username in to search for the ObjectID
$AzureUser = Get-AzureADUser -SearchString $User

#Pipe the ObjectID into the command to pull membership
Get-AzureADUserMembership -ObjectId $AzureUser.ObjectID | Export-Csv C:\Temp\$User.Csv

Read-Host -Prompt 'If there are no errors, the exported file is in the C:\temp folder. Press Enter to exit'
}
