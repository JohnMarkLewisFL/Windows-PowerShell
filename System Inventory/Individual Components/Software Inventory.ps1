# This script is designed to check the current software inventory of a PC and saves the results to a file


# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
Get-Module ImportExcel | Import-Module -Force # Put this towards the top of the script if you're planning on saving to Excel formats

# Creates a list of installed software along with the listed columns in the "Select-Object" section
Write-Host "Please wait while the list of installed software is being created. You will be prompted to save the list shortly."
$InstalledSoftwareList = Get-CimInstance -Class Win32_Product | Select-Object Name,Version,Vendor,InstallDate,Description

# Adds the Windows Form assembly type so the save dialog box will function properly
Add-Type -AssemblyName System.Windows.Forms

# Creates the save dialog box
$FileSaveDialogBox = New-Object System.Windows.Forms.SaveFileDialog
$FileSaveDialogBox.Filter = "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|Text (*.txt)|*.txt" # Sets the default file type options
$FileSaveDialogBox.Title = "Save As"
$FileSaveDialogBox.FileName = "Installed Software - $env:COMPUTERNAME"  # Sets the default file name (can be changed by the user)
#$FileSaveDialogBox.InitialDirectory = [Environment]::GetFolderPath('Desktop') # Sets the default folder location of the save file dialog box if desired

# Opens the save dialog box and sets the file type as noted by the user
$result = $FileSaveDialogBox.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFileType = [System.IO.Path]::GetExtension($FileSaveDialogBox.FileName)
    switch ($SelectedFileType) {
        ".csv" {
            $InstalledSoftwareList | Export-Csv -Path $FileSaveDialogBox.FileName -NoTypeInformation
        }
        ".xlsx" {
            $InstalledSoftwareList | Export-Excel -Path $FileSaveDialogBox.FileName
        }
        ".txt" {
            $InstalledSoftwareList | Out-File -FilePath $FileSaveDialogBox.FileName -Encoding UTF8
        }
    }
}