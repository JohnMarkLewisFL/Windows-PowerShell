# This script is designed to browse for a directory/folder, list its contents (recursively or not), and save this list as a file

# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
# If there are error messages regarding the 'Export-Excel' cmdlet, then run the following line in a separate PowerShell window to install the ImportExcel module manually:
# Install-Module -Name ImportExcel

Get-Module ImportExcel | Import-Module -Force 


# Needed to show the dialog boxes

Add-Type -AssemblyName System.Windows.Forms


# Opens the dialog box to select the directory/folder 

$FolderDialogBox = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderDialogBox.ShowDialog() | Out-Null
$FolderSelection = $FolderDialogBox.SelectedPath


# Retrieves the contents of the above-selected directory/folder

# Prompt the user for recursion in the results or not

$RecursionOption = Read-Host "Please enter 'yes' for recursion or 'no' for no recursion"


# Validates the user input

if ($RecursionOption -eq "yes") {
    Write-Host "The results will include recursion"
    # Note the inclusion of the -Recurse parameter and the inclusion of the -Name parameter to give much less verbose results
    $FolderContents = Get-ChildItem -Path $FolderSelection -Name -Recurse -Force
} elseif ($RecursionOption -eq "no") {
    Write-Host "The results will not include recursion"
    # Note the lack of the -Recurse parameter and the inclusion of the -Name parameter to give much less verbose results
    $FolderContents = Get-ChildItem -Path $FolderSelection -Name
} else {
    Write-Host "Please enter 'yes' or 'no'"
}


# Creates the save dialog box

$FileSaveDialogBox = New-Object System.Windows.Forms.SaveFileDialog
$FileSaveDialogBox.Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|Text (*.txt)|*.txt" # Sets the default file type options
$FileSaveDialogBox.Title = "Save As"
$FileSaveDialogBox.FileName = "Folder Contents"  # Sets the default file name (can be changed by the user)


# Opens the save dialog box and sets the file type as noted by the user

$result = $FileSaveDialogBox.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFileType = [System.IO.Path]::GetExtension($FileSaveDialogBox.FileName)
    
    switch ($SelectedFileType) {
        ".csv" {
            $FolderContents | Export-Csv -Path $FileSaveDialogBox.FileName -NoTypeInformation
        }
        ".xlsx" {
            $FolderContents | Export-Excel -Path $FileSaveDialogBox.FileName
        }
        ".txt" {
            $FolderContents | Out-File -FilePath $FileSaveDialogBox.FileName -Encoding UTF8
        }
    }
}