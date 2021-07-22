# Script to grab all members of a Dynamic DL and export the results to C:\temp\DDLMembership.csv
# Created by Griffin Rodgers

#Requires -RunAsAdministrator

Function Export-OnPremDDLMembership{

    CheckExchangeRemotelyRunning

    $DDL = Read-Host -Prompt "Please enter the name of the on prem Dynamic Distribution List you want to export the membership of"
    Write-Host "Please select the folder where you want the CSV to be saved"
    $Path = Get-FolderPath -Description "Select the folder where you want the CSV to be saved"

    if ([string]::IsNullOrEmpty($Path)) {throw "User closed the folder dialog. Exiting script."}

    $FTE = Get-DynamicDistributionGroup $DDL
    $OutputPath = $Path + "\$FTE.csv"
    Get-Recipient -RecipientPreviewFilter $FTE.RecipientFilter -ResultSize Unlimited -OrganizationalUnit $FTE.RecipientContainer | Select Name, Title, City, Department | Export-Csv $OutputPath -NoTypeInformation

    Read-Host -Prompt "The membership list has been exported to $OutputPath. Please press enter to close the script."
}
