# This script will check a Windows host for a list of its network adapters

# Checks if the ImportExcel module is installed
If (!(Get-Module -ListAvailable -Name ImportExcel)) {
    # Installs the ImportExcel module if it is not already installed
    Install-Module -Name ImportExcel -Force -AllowClobber
}

# Adds the Windows Form assembly type so the save dialog box will function properly
Add-Type -AssemblyName System.Windows.Forms

# Gets the computer's hostname
$Hostname = [System.Net.Dns]::GetHostName()
$NetworkAdapterList = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | 
                        ForEach-Object {
                            $ipv4 = $_.IPAddress | Where-Object { $_ -notmatch ":" }
                            $ipv6 = $_.IPAddress | Where-Object { $_ -match ":" }
                            $obj = New-Object PSObject -Property ([ordered]@{
                                Description = $_.Description
                                MACAddress = $_.MACAddress
                                IPAddressIPv4 = $ipv4
                                IPAddressIPv6 = $ipv6
                                IPSubnet = ($_.IPSubnet -join ", ")
                            })
                            $obj
                        }

# Builds the file name for the results/log
$FileName = "$Hostname - Network Adapter Inventory"

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
        1 { $NetworkAdapterList | Export-Excel -Path $SaveLocation -AutoSize }
        2 { $NetworkAdapterList | Export-Csv -Path $SaveLocation -NoTypeInformation }
        3 { $NetworkAdapterList | Out-File -FilePath $SaveLocation }
    }
}