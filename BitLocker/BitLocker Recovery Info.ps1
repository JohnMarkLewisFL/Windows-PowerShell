# This script is designed to retrieve BitLocker recovery keys for a user-specified drive letter


# Run as Administrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$args`"" -Verb RunAs; exit }


# Prompts the user for the drive letter and sanitizes the input
Do {
    $DriveLetter = (Read-Host "Please enter a drive letter to retrieve its BitLocker information").ToUpper()
    if ($DriveLetter -notmatch "^[A-Z]$") {
        Write-Host "Invalid input. Please enter a single letter from A to Z."
    }
} While ($DriveLetter -notmatch "^[A-Z]$")

 
 # Adds the Windows Form assembly type so the save dialog box will function properly
 Add-Type -AssemblyName System.Windows.Forms

 # Creates the save dialog box
 $FileSaveDialogBox = New-Object System.Windows.Forms.SaveFileDialog
 $FileSaveDialogBox.Filter = "Text (*.txt)|*.txt"
 $FileSaveDialogBox.Title = "Save BitLocker Info As"
 $FileSaveDialogBox.FileName = "BitLocker Information for drive $DriveLetter on $env:COMPUTERNAME" # Sets the default file name (can be changed by the user)
 #$FileSaveDialogBox.InitialDirectory = [Environment]::GetFolderPath('Desktop') # Sets the default folder location of the save file dialog box if desired

# Retrieves the BitLocker information
$BitLockerCommand = (Get-BitLockerVolume -MountPoint $DriveLetter).KeyProtector + $FileSaveDialogBox.ShowDialog()

# Runs the BitLocker information retrieval command and saves it to a .txt file
$BitLockerCommand | Out-File -FilePath $FileSaveDialogBox.FileName -Encoding UTF8