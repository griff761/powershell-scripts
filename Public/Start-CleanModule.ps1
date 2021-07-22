# Added to the Powershell Scripts Module by Griffin Rodgers
# Source of script is https://www.myerrorsandmysolutions.com/how-to-uninstall-older-versions-of-a-powershell-module-installed/

Function Start-CleanModule {

$ModuleName = 'PowershellScriptsModule'
$Latest = Get-InstalledModule $ModuleName 
Get-InstalledModule $ModuleName -AllVersions | ? {$_.Version -ne $Latest.Version} | Uninstall-Module
}