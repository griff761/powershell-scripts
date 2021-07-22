Function Get-FileName($InitialDirectory = [Environment]::GetFolderPath("Desktop"), $Title = "Select the CSV to be imported")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.Title = $Title
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

Function Get-FolderPath($InitialDirectory = [Environment]::GetFolderPath("Desktop"), $Description = "Select the output folder")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $OpenFolderDialog.RootFolder = 'MyComputer'
    $OpenFolderDialog.Description = $Description
    $OpenFolderDialog.ShowDialog() | Out-Null
    $OpenFolderDialog.SelectedPath
}
