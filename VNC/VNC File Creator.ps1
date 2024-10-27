# This script creates .vnc files from the user-provided list
# You may need to install the Import-Excel module for .xlsx compatibility
# If a password or encrypted password needs to be set, please see the $VNCFileContent variables towards the end of the script

# Adds the necessary assemblies for Windows Forms
# These are used for the file dialog box to select the URL list and the file dialog box to select a destination folder
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Introduction message
Write-Host "This script will create a .vnc file for each hostname or IP address in the list you specify shortly"
Start-Sleep -Seconds 3

# Prompts the user to enter a VNC port number or skip this step
Function Get-VNCPort {
    $VNCPort = Read-Host "`nPlease enter your VNC port number (press Enter to skip)"
    If ($VNCPort -eq "") {
        Write-Host "`nNo VNC port was entered"
    } Else {
        [int]$VNCPort = $VNCPort
        Write-Host "`nVNC port will be set to: $VNCPort"
    }
    Return $VNCPort
}

$VNCPort = Get-VNCPort

# Prompts the user to select the hosts list file
Write-Host "`nPlease select a file for a list of hostnames or IP addresses"
Write-Host "Please note that .csv and .xlsx files will need a column header named Host to function properly"
Start-Sleep -Seconds 3

# Function to open a file dialog box and select a URL list file
Function Select-URLListFile {
    $OpenFileDialogBox = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialogBox.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $OpenFileDialogBox.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx"
    $OpenFileDialogBox.Multiselect = $false
    $OpenFileDialogBoxResult = $OpenFileDialogBox.ShowDialog()
    If ($OpenFileDialogBoxResult -eq [System.Windows.Forms.DialogResult]::OK) {
        Return $OpenFileDialogBox.FileName
    }
    Return $Null
}

# Function to read hostnames/IP addresses from the hosts list file
Function Get-HostsFromFiles {
    Param ($FilePath)
    $HostsList = @()
    $Ext = [System.IO.Path]::GetExtension($FilePath)
    Switch ($Ext) {
        ".txt" {
            $HostsList = Get-Content $FilePath
        }
        ".csv" {
            $HostsList = Import-Csv $FilePath | ForEach-Object { $_.Host }
        }
        ".xlsx" {
            $HostsList = Import-Excel $FilePath | ForEach-Object { $_.Host }
        }
    }
    Return $HostsList
}

# Start of the hosts list handling
$HostsListFile = Select-URLListFile
If (-Not $HostsListFile) {
    Write-Host "`nNo hosts list file was selected"
    Start-Sleep -Seconds 3
    Exit
}

# Path to the folder to save the .vnc files
Write-Host "`nPlease select a folder to save the .vnc files"
Start-Sleep -Seconds 3
$FolderDialogBox = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderDialogBox.ShowDialog() | Out-Null
$VNCFilesPath = $FolderDialogBox.SelectedPath

If (-Not $VNCFilesPath) {
    Write-Host "`nNo folder was selected"
    Start-Sleep -Seconds 3
    Exit
}

# Create .vnc files for each host
$Hosts = Get-HostsFromFiles $HostsListFile
ForEach ($Hostname in $Hosts) {
    If ($VNCPort -eq "") {
        # Use the top line with your encrypted VNC password if you wish to automatically set a VNC password in each .vnc file
        # Use # to comment out the $VNCFileContent line you don't need
        #$VNCFileContent = "ConnMethod=tcp`nHost=" + $Hostname + "`nPassword=InsertYourPassword"
        $VNCFileContent = "ConnMethod=tcp`nHost=" + $Hostname
    } Else {
        # Use the top line with your encrypted VNC password if you wish to automatically set a VNC password in each .vnc file
        # Use # to comment out the $VNCFileContent line you don't need
        #$VNCFileContent = "ConnMethod=tcp`nHost=" + $Hostname + ":" + $VNCPort + "`nPassword=InsertYourPassword"
        $VNCFileContent = "ConnMethod=tcp`nHost=" + $Hostname + ":" + $VNCPort
    }
    $FileName = [System.IO.Path]::Combine($VNCFilesPath, $Hostname + ".vnc")
    New-Item -Path $FileName -ItemType File -Value $VNCFileContent
}

Write-Host "`nThe list of .vnc files has been created in $VNCFilesPath"
Start-Sleep -Seconds 3
Pause