# This script will download the favicon.ico of specified URLs. There is an option to enter a single URL or use a list of URLs.

# If using a list file a log file will be saved in the same location as the .ico files.
# If using a .xlsx file, cell A1 will be ignored, so put a header or column name like "URL" to use all listed URLs.
# The format of the .txt, .csv, and .xlsx files should have one URL per line. Headers such as http:// and https:// will be added automatically.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function Download-Favicon($URL, $SavePath, $LogFile) {
    # Check if the URL already starts with 'http://' or 'https://'
    If (-not ($URL.StartsWith('http://') -or $URL.StartsWith('https://'))) {
        $URL = 'http://' + $URL
    }

    # Extract the domain name from the URL
    $URI = New-Object System.URI($URL)
    $Domain = $URI.Host.Split('.')[0]

    # Download the favicon.ico file
    # If this errors out, check the URL manually to see if the favicon.ico file exists or if it may be in a different format (such as .gif or .png)
    $FaviconURL = $URL + '/favicon.ico'
    $WebClient = New-Object System.Net.WebClient
    Try {
        $WebClient.DownloadFile($FaviconURL, $SavePath)
        If ($LogFile) {
            Add-Content -Path $LogFile -Value "Success: Downloaded favicon from $FaviconURL"
        }
    }
    Catch {
        Write-Host "Failed to download the favicon from $FaviconURL"
        If ($LogFile) {
            Add-Content -Path $LogFile -Value "Error: Failed to download favicon from $FaviconURL"
        }
        Start-Sleep -Seconds 3
    }
}

# Ask the user if they want to enter a single URL or use a .txt file
$Choice = Read-Host -Prompt 'Type 1 to enter a single URL manually
Type 2 to select a list of multiple URLs in a .txt, .csv, or .xlsx file (with logging)'

if ($Choice -eq 1) {
    do {
        # Prompt the user to enter a URL
        $URL = Read-Host -Prompt 'Enter the URL'
        $URI = New-Object System.URI('http://' + $URL)
        $D = $URI.Host.Split('.')[0]

        # Create a new SaveFileDialog object
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.filter = "ICO files (*.ico)| *.ico"
        $SaveFileDialog.FileName = $D + ".ico"
        $SaveFileDialog.Title = "Save an Image File"
        $SaveFileDialog.ShowDialog() | Out-Null

        If ($SaveFileDialog.FileName) {
            Download-Favicon $URL $SaveFileDialog.FileName $null
        }
        Else {
            Write-Host "No file name specified, exiting."
        }

        # Ask the user if they want to enter another URL
        $choicC = Read-Host -Prompt 'Enter 1 to enter another URL or 2 to exit'
    } While ($choicC -eq 1)
}
ElseIf ($Choice -eq 2) {
    # Create a new OpenFileDialog object
    Write-Host "`nSelect a URL list in .txt, .csv, .xlsx, or .xls format (.xlsx and .xls will need cell A1 to be a heading such as URL)"
    Start-Sleep -Seconds 3
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "Text files (*.txt)| *.txt|CSV files (*.csv)| *.csv|Excel files (*.xlsx, *.xls)| *.xlsx, *.xls"
    $OpenFileDialog.Title = "Select a List File"
    $OpenFileDialog.ShowDialog() | Out-Null

    # Create a new FolderBrowserDialog object
    Write-Host "`nSelect a folder to save the .ico files and the .txt log file"
    Start-Sleep -Seconds 3
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.Description = "Select a Folder"
    $FolderBrowserDialog.ShowDialog() | Out-Null

    If ($OpenFileDialog.FileName -and $FolderBrowserDialog.SelectedPath) {
        # Determine the file type
        $FileType = [System.IO.Path]::GetExtension($OpenFileDialog.FileName)

        # Read the URLs from the file
        If ($FileType -eq ".txt") {
            $URLs = Get-Content $OpenFileDialog.FileName
        }
        ElseIf ($FileType -eq ".csv") {
            $URLs = Import-Csv $OpenFileDialog.FileName -Header URL | ForEach-Object { $_.URL }
        }
        ElseIf ($FileType -eq ".xlsx") {
            $URLs = Import-Excel $OpenFileDialog.FileName | ForEach-Object { $_.PSObject.Properties.Value }
        }

        $LogFile = Join-Path $FolderBrowserDialog.SelectedPath ("Favicon Harvester Log - " + (Get-Date -Format "yyyyMMddTHHmmss") + ".txt")

        ForEach ($URL in $URLs) {
            # Trim the URL to remove any leading or trailing whitespace
            $URL = $URL.Trim()

            # Check if the URL already starts with 'http://' or 'https://'
            If (-not ($URL.StartsWith('http://') -or $URL.StartsWith('https://'))) {
                $URL = 'http://' + $URL
            }

            # Extract the domain name from the URL
            $URI = New-Object System.URI($URL)
            $D = $URI.Host.Split('.')[0]

            # Generate the file name
            $FileName = Join-Path $FolderBrowserDialog.SelectedPath ($D + ".ico")

            Download-Favicon $URL $FileName $LogFile
        }
    }
    Else {
        Write-Host "No file or folder was selected, exiting"
    }
}
Else {
    Write-Host "Invalid choice, exiting."
}