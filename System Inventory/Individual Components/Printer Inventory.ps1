# This script will check a Windows host for all of its installed printers, note the printer name and printer port (i.e. IP or USB), and export it to a file


# Checks if the ImportExcel module is installed
If (!(Get-Module -ListAvailable -Name ImportExcel)) {
    # Installs the ImportExcel module if it is not already installed
    Install-Module -Name ImportExcel -Force -AllowClobber
}

# Adds the Windows Form assembly type so the save dialog box will function properly
Add-Type -AssemblyName System.Windows.Forms

# Gets the computer's hostname
$Hostname = [System.Net.Dns]::GetHostName()

# Gets all installed printers and their port names (software printers may show a strange port)
$Printers = Get-Printer | Select-Object Name, PortName

# Builds the file name for the results/log
$FileName = "$Hostname - Printer Inventory"

# Creates the SaveFileDialog box
$SaveFileDialogBox = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialogBox.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$SaveFileDialogBox.Filter = 'Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|Text file (*.txt)|*.txt'
$SaveFileDialogBox.FileName = $FileName

# Shows the SaveFileDialog box
$Result = $SaveFileDialogBox.ShowDialog()

# If the user clicked OK, save the file
If ($Result -eq 'OK') {
    $SaveLocation = $SaveFileDialogBox.FileName

    # Determine the file extension
    Switch ($SaveFileDialogBox.FilterIndex) {
        1 { $Printers | Export-Excel -Path $SaveLocation }
        2 { $Printers | Export-Csv -Path $SaveLocation -NoTypeInformation }
        3 { $Printers | Out-File -FilePath $SaveLocation }
    }
}