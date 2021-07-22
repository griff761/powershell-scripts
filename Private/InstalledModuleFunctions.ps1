Function InstallAzureAD {
    if (Get-Module -ListAvailable -Name "AzureAD") {
        Write-Host "AzureAD Module already installed"
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
}

Function InstallExchangeOnline {
    if (Get-Module -ListAvailable -Name "ExchangeOnlineManagement") {
        Write-Host "Exchange Online Module already installed"
    }
    else {
        try {
            Install-Module ExchangeOnlineManagement -Repository PSGallery
        }
        catch [Exception] {
            $_.message
            exit
        }
    }
}

Function InstallImportExcel {
    if (Get-Module -ListAvailable -Name "ImportExcel") {
        Write-Host "Import Excel Module already installed"
    }
    else {
        try {
            Install-Module ImportExcel -Repository PSGallery
        }
        catch [Exception] {
            $_.message
            exit
        }
    }
}

Function CheckAzureADRunning {
    #Check if user is already connected to AzureAD
    try 
    { $var = Get-AzureADTenantDetail } 

    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] 
    { Write-Host "You're not connected to AzureAD, connecting."; Connect-AzureAD }
}

Function CheckExchangeOnlineRunning {
    #Connect & Login to ExchangeOnline (MFA)
    try 
    { $var = Get-CloudAcceptedDomain } 

    catch [System.Management.Automation.CommandNotFoundException] 
    { Write-Host "You're not connected to Exchange Online, connecting."; Connect-ExchangeOnline -prefix Cloud }
}

Function CheckExchangeRemotelyRunning{
    $connected = Get-Variable ExchangeSession -ErrorAction SilentlyContinue

    if ($null -eq $connected) {
        Connect-ExchangeRemotely
    }
}
