# This script will gather all saved Wi-Fi SSIDs along with their respective PSKs, then save the name and file type as prompted by the user


# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
# If there are error messages regarding the 'Export-Excel' cmdlet, then run the following line in a separate PowerShell window to install the ImportExcel module manually:
# Install-Module -Name ImportExcel
Get-Module ImportExcel | Import-Module -Force

# Gather the SSID from the netsh command and split the unused text and whitespaces
$WirelessProfiles = netsh wlan show profiles | Select-String -Pattern "All User Profile" | ForEach-Object{ ($_ -split ":")[-1].Trim() };

# For each profile found with the netsh command, show the password in plaintext and display the SSID in columns
$WirelessProfiles | ForEach-Object {
	$ProfileData = netsh wlan show profiles name=$_ key="clear";
	$SSID = $ProfileData | Select-String -Pattern "SSID Name" | ForEach-Object{ ($_ -split ":")[-1].Trim().Trim('"') };
	$PSK = $ProfileData | Select-String -Pattern "Key Content" | ForEach-Object{ ($_ -split ":")[-1].Trim() };

    $List =	[PSCustomObject]@{
		SSID = $SSID;
		Password = $PSK
	    } 
        # Adds the Windows Form assembly type so the save dialog box will function properly
        Add-Type -AssemblyName System.Windows.Forms
        
        # Creates the save dialog box
        $FileSaveDialogBox = New-Object System.Windows.Forms.SaveFileDialog
        $FileSaveDialogBox.Filter = "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|Text (*.txt)|*.txt" # Sets the default file type options
        $FileSaveDialogBox.Title = "Save Wi-Fi Credentials As"
        $FileSaveDialogBox.FileName = "Wi-Fi Passwords - $env:computername"  # Sets the default file name (can be changed by the user)
        #$FileSaveDialogBox.InitialDirectory = [Environment]::GetFolderPath('Downloads') # Sets the default folder location of the save file dialog box if desired
        
        # Opens the save dialog box and sets the file type as noted by the user
        $result = $FileSaveDialogBox.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $SelectedFileType = [System.IO.Path]::GetExtension($FileSaveDialogBox.FileName)
            switch ($SelectedFileType) {
                ".csv" {
                    $List | Export-Csv -Path $FileSaveDialogBox.FileName -NoTypeInformation
                }
                ".xlsx" {
                    $List | Export-Excel -Path $FileSaveDialogBox.FileName
                }
                ".txt" {
                    $List | Out-File -FilePath $FileSaveDialogBox.FileName -Encoding UTF8
                }
            }
        }
    }