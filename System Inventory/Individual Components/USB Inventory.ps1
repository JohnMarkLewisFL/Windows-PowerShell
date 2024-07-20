# This script will check a Windows host for a list of its USB adapters and devices while filtering out most of the generic USB device and hub entries


# Checks if the ImportExcel module is installed
If (!(Get-Module -ListAvailable -Name ImportExcel)) {
    # Installs the ImportExcel module if it is not already installed
    Install-Module -Name ImportExcel -Force -AllowClobber
}

# Adds the Windows Form assembly type so the save dialog box will function properly
Add-Type -AssemblyName System.Windows.Forms

# Gets the computer's hostname
$Hostname = [System.Net.Dns]::GetHostName()

# Gets installed USB adapters and devices while filtering out most of the generic entries
$ExcludeUSBNames = "USB Composite Device", "USB Printing Support", "USB Mass Storage Device", "Generic SuperSpeed USB Hub", "USB Root Hub", "USB Root Hub (USB 3.0)", "Generic USB Hub"
$USBList = Get-PnpDevice | Where-Object { $_.Class -eq "USB" -and $ExcludeUSBNames -notcontains $_.Name } | Select-Object Name, Status

# Builds the file name for the results/log
$FileName = "$Hostname - USB Inventory"

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
        1 { $USBList | Export-Excel -Path $SaveLocation -AutoSize }
        2 { $USBList | Export-Csv -Path $SaveLocation -NoTypeInformation }
        3 { $USBList | Out-File -FilePath $SaveLocation }
    }
}