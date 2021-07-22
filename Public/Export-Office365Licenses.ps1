#This script can be used to export a CSV containing every user licensed for Office 365
#Created by Griffin Rodgers

#Requires -RunAsAdministrator

Function Export-Office365Licenses{

Write-Host "Select the folder where you want the licenses CSV to be saved" -ForegroundColor Green

$Path = Get-FolderPath -Description "Select the folder where you want the licenses CSV to be saved"
if ([string]::IsNullOrEmpty($Path)) {throw "User closed the folder dialog. Exiting script."}

$Location = $Path + '\Office365LicensedUsers.csv '

Get-ADGroupMember -Identity 'Office 365 All Apps' <#This group name can be changed to whatever AD group assigns O365 licenses#> | Get-ADUser -Property Name, userPrincipalName, employeeNumber, lastlogonDate | Select Name, userPrincipalName, employeeNumber, lastlogonDate | Export-Csv $Location -NoTypeInformation
}
