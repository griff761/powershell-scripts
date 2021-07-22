#This script will pull all the Azure AD and AD groups a user belongs to and export it to an Excel spreadsheet
#Created by Griffin Rodgers
#Requires -RunAsAdministrator

Function Export-UserGroupMembership {

InstallAzureAD

CheckAzureADRunning

#Get the user's account information
$User = Read-Host -Prompt 'Enter the username for the account'

#Pipe the username in to search for the ObjectID
$AzureUser = Get-AzureADUser -SearchString $User

Write-Host "Please select the folder where you want the user's group membership CSV to be saved"
$Path = Get-FolderPath -Description "Select the folder where you want the user's group membership CSV to be saved"
if ([string]::IsNullOrEmpty($Path)) {throw "User closed the folder dialog. Exiting script."}

$OutputPath = $Path + "\$User.csv"

#Pipe the ObjectID into the command to pull membership
Get-AzureADUserMembership -ObjectId $AzureUser.ObjectID | Export-Csv $OutputPath -NoTypeInformation

Read-Host -Prompt "If there are no errors, the exported file is at $OutputPath. Press Enter to exit"
}
