Function getExpiringUsers {
    #This script searches for accounts which will expire soon
    $filter = "(&(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)(accountExpires>=1)(accountExpires<=" + (get-date).adddays(30).ToFileTimeUtc() + "))" 

    $results = Get-ADUser -SearchBase "OU=Users" -ldapFilter $filter -Properties Name, description, accountExpires | Select Name, description, accountExpires | Sort-Object accountExpires, Name #correct SearchBase required

    # Using $results.Name.Length works around $results.Count not working if only one result
    if ($results.Name.Length -gt 0) {
        Write-Output "User accounts which will expire in the next 30 days"
        Write-Output "Review these accounts with HR and the user's manager"
        #Write-Output "LDAP filter used: $filter"
        foreach ($result in $results) {
            $accountExpires = [datetime]::FromFileTimeUTC($result.accountExpires[0])
            #$accountExpiresDiff = new-TimeSpan $accountExpires $(Get-Date);
            Write-Output "`t$($result.Name) (expires $($accountExpires.toShortDateString()))" 
        }
        Write-Output "`n`n"
    }
}

Function getObjectsOwnedByDisabledUser {
    #This script searches for objects managed by a disabled user
    # Note: you cannot filter Get-ADUser for distinguishedName

    $results = @(Get-ADObject -LDAPFilter "(|(managedBy=*)(manager=*))" -Properties manager, managedBy | `
            Where-Object { $_.managedBy -like "*Disabled*" -or $_.manager -like "*Disabled*" } | `
            Select-Object name, objectclass, @{label = 'Manager'; expression = { $_.manager -replace '^CN=|,.*$' } }, @{label = 'ManagedBy'; expression = { $_.managedBy -replace '^CN=|,.*$' } } | `
            Sort-Object -Property ManagedBy, Manager, objectClass, name)

    #write-output $results

    if ($results.Count -gt 0) {
        Write-Output "AD objects managed by a disabled user"
        Write-Output "Identify and set a new manager"
        foreach ($result in $results) {
            Write-Output "`t$($result.Manager)$($result.ManagedBy): $($result.Name) ($($result.objectClass)) " 
        }
        Write-Output "`n`n"
    }

}

Function getUsersInOnboardingGroupOver30 {
    #This script searches for users in the Newly Onbarded Users group who are more than 30 days old

    $filter = "(&(memberOf=CN=Newly Onboarded Users,OU=Groups(whenCreated<=" + (get-date).adddays(-30).ToUniversalTime().ToString('yyyyMMddHHmmss.0Z') + "))" #Correct searchbase required

    $results = @(Get-ADUser -LDAPFilter $filter | Sort-Object Name)

    if ($results.Count -gt 0) {
        Write-Output "$($results.Count) Account(s) more than 30 days old in the Newly Onboarded Users group"
        Write-Output "Remove these users from the Newly Onboarded Users AD group"
        foreach ($result in $results) {
            Write-Output "`t$($result.Name)" 
        }
        Write-Output "`n`n"
    }
}

Function getUsersWithExpiredPW {
    #This script searches for accounts which have expired but are not disabled
    $filter = "(&(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)" + `
        "(msDS-LastSuccessfulInteractiveLogonTime>=1)(msDS-LastSuccessfulInteractiveLogonTime<=" + (get-date).adddays(-90).ToFileTimeUtc() + "))" 

    $results = Get-ADUser -SearchBase "OU=Users" `
        -ldapFilter $filter `
        -Properties Name, description, pwdLastSet, 'msDS-LastSuccessfulInteractiveLogonTime' | `
        Sort-Object Name | `
        Where-Object {
        # Don't include service accounts
        $_.DistinguishedName -notlike "CN=*,OU=Service-Accts,OU=Users" -And
        $_.DistinguishedName -notlike "CN=*,OU=Service-Accts-AAD,OU=Users"  
    } #Correct SearchBase required

    # Using $results.Name.Length works around $results.Count not working if only one result
    if ($results.Name.Length -gt 0) {
        Write-Output "User accounts which are expired"
        Write-Output "Disable these accounts and process as a termination"
        #Write-Output "LDAP filter used: $filter"
        foreach ($result in $results) {
            if ($result.'msDS-LastSuccessfulInteractiveLogonTime' -ne $null) {
                $lastLogon = [datetime]::FromFileTimeUTC($result.'msDS-LastSuccessfulInteractiveLogonTime')
                $Logondatediff = new-TimeSpan $lastLogon $(Get-Date);
            }

            if ($result.pwdLastSet -ne $null) {
                $pwdLastSet = [datetime]::FromFileTimeUTC($result.pwdLastSet[0])
                $pwdLastSetDiff = new-TimeSpan $pwdLastSet $(Get-Date);
            }

            Write-Output "`t$($result.Name) (Last Logon: $($Logondatediff.Days) days ago): Password age: $($pwdLastSetDiff.Days) days"
        }
        Write-Output "`n`n"
    }

}
