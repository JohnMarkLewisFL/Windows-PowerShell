# This script will check a Windows host for a list of its Bluetooth adapters and Bluetooth devices


# Checks if the ImportExcel module is installed
If (!(Get-Module -ListAvailable -Name ImportExcel)) {
    # Installs the ImportExcel module if it is not already installed
    Install-Module -Name ImportExcel -Force -AllowClobber
}

# Adds the Windows Form assembly type so the save dialog box will function properly
Add-Type -AssemblyName System.Windows.Forms

# Gets the computer's hostname
$Hostname = [System.Net.Dns]::GetHostName()

# Gets all installed Bluetooth adapters and devices
$BluetoothList = Get-PnpDevice | Where-Object { $_.Class -eq "Bluetooth" } | Select-Object Name, Status

# Builds the file name for the results/log
$FileName = "$Hostname - Bluetooth Inventory"

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
        1 { $BluetoothList | Export-Excel -Path $SaveLocation -AutoSize }
        2 { $BluetoothList | Export-Csv -Path $SaveLocation -NoTypeInformation }
        3 { $BluetoothList | Out-File -FilePath $SaveLocation }
    }
}