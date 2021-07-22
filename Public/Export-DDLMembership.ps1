# Script to grab all members of a Dynamic DL and export the results to C:\temp\DDLMembership.csv
# Created by Griffin Rodgers

#Requires -RunAsAdministrator
function Export-DDLMembership {
    InstallExchangeOnline

    CheckExchangeOnlineRunning

    $DDL = Read-Host -Prompt "Please enter the name of the cloud Dynamic Distribution List you want to export the membership of"

    $FTE = Get-CloudDynamicDistributionGroup $DDL
    Get-CloudRecipient -RecipientPreviewFilter $FTE.RecipientFilter -OrganizationalUnit $FTE.RecipientContainer | Select Name, Title, City, Department | Export-Csv C:\temp\$FTE.csv

    Read-Host -Prompt "The membership list has been exported to C:\temp\$FTE.csv. Please press enter to close the script."
}
