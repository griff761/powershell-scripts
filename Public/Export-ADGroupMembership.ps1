#This script can be used to pull group membership from a group in AD or AzureAD
#Created by Griffin Rodgers

#Requires -RunAsAdministrator

function Export-ADGroupMembership {

    InstallAzureAD
    
    CheckAzureADRunning

    $Group = Read-Host -Prompt 'Enter the distribution or security group name'

    $ObjectID = Get-AzureADGroup -SearchString $Group

    Write-Host "Please select the folder where you want the group membership CSV to be saved"
    $Path = Get-FolderPath -Description "Select the folder where you want the group membership CSV to be saved"

    if ([string]::IsNullOrEmpty($Path)) {throw "User closed the folder dialog. Exiting script."}

    $OutputPath = $Path + "\$Group.csv"

    Get-AzureADGroupMember -ObjectID $ObjectID.ObjectId | Select DisplayName, UserPrincipalName, ObjectType | Export-Csv $OutputPath -NoTypeInformation

    Read-Host -Prompt "If there are no errors, the CSV was exported to $OutputPath. Press Enter to exit"
}
