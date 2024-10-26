# This script will automatically open URLs from a user-provided list
# This can be helpful if dozens (or hundreds or thousands) of URLs need to be checked
# Please note the Import-Excel module will need to be installed for proper functionality

# Adds the necessary assemblies for Windows Forms
# These are used for the file dialog box to select the URL list
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Prompts the user to select a web browser from the following list
# If the user selects a web browser that is not currently installed on their computer, then the script will fail
# This list and the validation switch statement can easily be expanded or trimmed as needed
Function Select-WebBrowser {
    Do {
        Write-Host "1. Chrome"
        Write-Host "2. Edge"
        Write-Host "3. Firefox"
        $UserWebBrowserSelection = Read-Host "`nPlease enter the number of the browser you want to use"

        # Validates the user's input and prompts to re-enter a valid selection
        Switch ($UserWebBrowserSelection) {
            1 { $WebBrowser = "chrome"; $Valid = $True }
            2 { $WebBrowser = "msedge"; $Valid = $True }
            3 { $WebBrowser = "firefox"; $Valid = $True }
            Default {
                Write-Host "Invalid selection. Please enter 1, 2, or 3."
                $Valid = $False
            }
        }
    } While (-Not $Valid)
    
    Return $WebBrowser
}

# Sets the WebBrowser variable to the user's selection
$WebBrowser = Select-WebBrowser

# Function to open a file dialog box and select a URL list file
Write-Host "`nPlease select a URL list file"
Write-Host "You can select a .txt file, a .csv, or a .xlsx file"
Write-Host "Please note the .csv and .xlsx files will need a column header named URL to properly retrieve the URLs"
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

# Function to read URLs from files
Function Get-URLsFromFiles {
    Param ($FilePath)
    $URLs = @()
    $Ext = [System.IO.Path]::GetExtension($FilePath)
    Switch ($Ext) {
        ".txt" {
            $URLs = Get-Content $FilePath
        }
        ".csv" {
            $URLs = Import-Csv $FilePath | ForEach-Object { $_.URL }
        }
        ".xlsx" {
            $URLs = Import-Excel $FilePath | ForEach-Object { $_.URL }
        }
    }
    Return $URLs
}

# Start of the URL list handling
$URLListFile = Select-URLListFile
If (-not $URLListFile) {
    Write-Host "`nNo URL list file was selected"
    Start-Sleep -Seconds 3
    Exit
}

# Verification before opening the first batch of URLs
Write-Host "`nYou have selected: $URLListFile and the web browser $WebBrowser"
Start-Sleep -Seconds 3
Write-Host "`nThe first batch of URLs will open shortly"
Write-Host "`nAfter each batch of URLs, please go to this PowerShell window and press the Enter key to proceed"
Start-Sleep -Seconds 3

# Define the URL list
$URLList = Get-URLsFromFiles $URLListFile

# Automatically opens each URL in a web browser tab
$URLCount = 0
# Change the number in $NumberOfTabs to the number of tabs you want to open in each batch
$NumberOfTabs = 25
ForEach ($URL in $URLList) {
    $BrowserLaunch = Start-Process $WebBrowser $URL
    # Pauses the script after the set number of tabs, resumes when the enter key is pressed
    If (++$URLCount % $NumberOfTabs -eq 0) {
        Pause
    }
    $BrowserLaunch
}

# Ending message
Write-Host "`nAll URLs in $URLListFile have been opened with $WebBrowser"
Start-Sleep -Seconds 3
Write-Host "`nThis PowerShell window will close momentarily"
Start-Sleep -Seconds 5