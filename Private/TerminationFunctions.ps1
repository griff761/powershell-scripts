Function SaveTerminatedUserGroups($ADUser, $Path = "C:\temp\") {

    $Groups = Get-ADPrincipalGroupMembership $ADUser | Where-Object -Property Name -Ne -Value 'Domain Users'
    $TodayDash = (Get-Date).ToString('MM-dd-yy')
    $OutputPath = $Path + $ADUser.Name + '_' + $TodayDash + '.xlsx'

    $ADUser | Select cn, Description, DistinguishedName, whenCreated |Export-Excel -Path $OutputPath -WorksheetName "Account Info" -ClearSheet -AutoSize 
    $excelPackage = Open-ExcelPackage -Path $OutputPath 
    $Groups | Select distinguishedName, GroupCategory, GroupScope, name | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Groups" -AutoSize -ClearSheet

    Write-Host -ForegroundColor Green "File containing user groups saved to $OutputPath."
}
