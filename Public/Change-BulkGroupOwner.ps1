# User inputs a CSV file containg a Groups column and a New Owner column and the script will change the owner of the group to the person in New Owner
# Written by Griffin Rodgers

#Requires -RunAsAdministrator

Function Change-BulkGroupOwner{

Write-Host "Please select the file containing groups and new owner (headers must be Group and New Owner)" -ForegroundColor Green
$Path = Get-FileName -Title "Select the file containing groups and new owner (headers must be Group and New Owner)"
if ([string]::IsNullOrEmpty($Path)) {throw "Choose file dialog closed. Exiting script."}

$Ticket = Read-Host "Please enter the ticket number for this bulk group ownership change"
$Initials = Read-Host "Please enter your initials"

$Groups = Import-CSV $Path

if ([string]::IsNullOrEmpty($Groups[0].Group) -or [string]::IsNullOrEmpty($Groups[0].'New Owner')) {throw "Incorrect CSV headers"}

ForEach ($Group in $Groups)
{
    $Name = $Group.'New Owner' + "*"
    $NewOwner = Get-ADUser -Filter {(name -like $Name) -and (name -notlike "* - ADM")}
    $Desc = "New Owner: $Name`r`nAs per $Ticket $Initials"
    $ADGroup = Get-ADGroup -Identity $Group.Group -Properties info
    Set-ADGroup $ADGroup -ManagedBy $NewOwner -Replace @{info="$Desc `r`n$($ADGroup.info)"}
}
}