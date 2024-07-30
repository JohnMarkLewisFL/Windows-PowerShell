# This script is used to bulk rename file extensions efficiently and log the before and after states of the affected folders/directories. 
# It offers an option for recursion. The file extensions are case-sensitive.


# Adds the Windows Forms assembly for creating the dialog boxes
Add-Type -AssemblyName System.Windows.Forms

# Initial message to the user
Write-Host "This script will perform a bulk rename of file extensions. Please note it is case-sensitive.`n"
Start-Sleep -Seconds 2
Write-Host "This script logs the before and after states of the affected folders. You will now be prompted to select a save location for the .txt log file.`n"
Start-Sleep -Seconds 2

# Creates the file save dialog box for the log file
$LogSaveDialogBox = New-Object System.Windows.Forms.SaveFileDialog
$LogSaveDialogBox.Filter = "Text files (*.txt)|*.txt"
$LogSaveDialogBox.FileName = "Extension Rename Results - $(Get-Date -Format 'yyyyMMddHHmmss').txt"
$LogSaveDialogBox.Title = "Save Log File As"

# Shows the file save dialog box for the log file and retrieves its path
$LogSaveDialogBox.ShowDialog() | Out-Null
$LogFilePath = $LogSaveDialogBox.FileName

# Asks the user if they want to use recursion or not and performs input validation
$UseRecursion = $Null
While ($Null -eq $UseRecursion) {
    $Input = Read-Host -Prompt 'Do you want to use recursion, which will affect files in subfolders? (yes/no)'
    $Input = $Input.ToLower()

    If ($Input -eq 'yes' -or $Input -eq 'y') {
        $UseRecursion = $True
    } ElseIf ($Input -eq 'no' -or $Input -eq 'n') {
        $UseRecursion = $False
    } Else {
        Write-Host `n"Invalid input. Please enter 'yes' or 'no'."
    }
}
$Recursive = $UseRecursion

# Creates the dialog box to browse for a folder
$BrowseForFolder = New-Object System.Windows.Forms.FolderBrowserDialog
$BrowseForFolder.ShowDialog() | Out-Null
$FolderPath = $BrowseForFolder.SelectedPath

# Gets all unique file extensions in the user-specified folder (and subfolders if recursion is selected)
$FileExtensions = Get-ChildItem -Path $FolderPath -File -Recurse:$Recursive | Select-Object -ExpandProperty Extension -Unique

# Displays the unique extensions with each one on a new line
Write-Host "`nThe unique file extensions in ${FolderPath} are:"
$FileExtensions | ForEach-Object {
    Write-Host $_
}

# Logs the initial state of the files
"Initial state of files:" | Out-File -FilePath $LogFilePath
Get-ChildItem -Path $FolderPath -Recurse:$Recursive | Out-File -FilePath $LogFilePath -Append

# Prompt the user to enter a currently existing file extension in that selected folder
$OldExtension = Read-Host -Prompt `n'Please enter the current file extension (case-sensitive) you wish to change, including the dot'

# Check if the old extension exists in the folder
While ($OldExtension -notin $FileExtensions) {
    Write-Host `n"No files with the extension *$OldExtension found in ${FolderPath}."
    $OldExtension = Read-Host -Prompt `n'Please enter a valid file extension'
}

# Prompt the user to provide another file extension to bulk rename them
$NewExtension = Read-Host -Prompt `n'Please enter the new file extension (case-sensitive), including the dot'

# Tells the user the files will be renamed and sleeps for 3 seconds
Write-Host `n"Files with the extension $OldExtension will be renamed to $NewExtension"`n
Start-Sleep -Seconds 3

# Get all files with the exact old extension in the selected folder
$Files = Get-ChildItem -Path $FolderPath -Recurse:$Recursive | Where-Object { $_.Extension -eq $OldExtension }

# Loops through each file and renames it
ForEach ($File In $Files) {
    # Get the base name of the file (without extension)
    $BaseName= [System.IO.Path]::GetFileNameWithoutExtension($File)

    # Constructs the new file name
    $NewFileName = $BaseName + $NewExtension

    # Renames the file
    Rename-Item -Path $File.FullName -NewName $NewFileName
}

# Logs the final state of the files
"`n`nFinal state of files:" | Out-File -FilePath $LogFilePath -Append
Get-ChildItem -Path $FolderPath -Recurse:$Recursive | Out-File -FilePath $LogFilePath -Append

Write-Host `n"All *$OldExtension files in ${FolderPath} have been renamed to *$NewExtension."
Write-Host `n"The log file (which notes the original and final states of all files affected) is saved at $LogFilePath"
Start-Sleep -Seconds 5