# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# This script uses the NVD (National Vulnerability Database) to check for the top 10 CVEs using the CIRCL CVE Search API

# Introductory message
Write-Host "This script uses the CIRCL CVE Search API to retrieve a list of the top CVEs in the NVD.`nAn active internet connection is required.`n`n"

# Prompt the user to input the number of CVEs to view (1 to 30) and validate the input
Do {
    $CVEAmount = Read-Host "Please enter the number of top CVEs you wish to view (1 to 30):`n"
} While ($CVEAmount -lt 1 -or $CVEAmount -gt 30 -or $CVEAmount -match "\D")

# Ask the user if they want to save the CVE list to a .txt file and validate the input
Do {
    $SaveToFile = (Read-Host "`nDo you want to save the CVE list to a .txt file? (yes/no):`n").ToLower()
} While ($SaveToFile -ne "yes" -and $SaveToFile -ne "no" -and $SaveToFile -ne "y" -and $SaveToFile -ne "n")

# If the user wants to save the CVE list to a .txt file, show the save file dialog box
if ($SaveToFile -eq "yes" -or $SaveToFile -eq "y") {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $SaveFileDialog.Filter = "Text files (*.txt)|*.txt"
    $SaveFileDialog.FileName = "Top $CVEAmount CVEs - $(Get-Date -Format 'yyyyMMddHHmmss').txt"
    $SaveFileDialog.Title = "Save CVE List"
    $Result = $saveFileDialog.ShowDialog()

    # If the user clicks the Save button in the dialog, save the file path
    If ($Result -eq "OK") {
        $FilePath = $SaveFileDialog.FileName
    }
}

# Define the CIRCL CVE Search API
$URI = 'https://cve.circl.lu/api/last'

# Send the API request and get the response
$Response = Invoke-RestMethod -Uri $URI -Method Get

# Parse the response to get the CVE data
$CVEs = $Response

# Select the top 10 CVEs from the response
$TopCVEs = $CVEs | Select-Object -First $CVEAmount

# Initialize a counter
$Counter = 1

# Extra line for formatting
Write-Host `n

# Loop through each CVE and print its ID and summary along with some formatting for readability
ForEach ($CVE in $TopCVEs) {
    $CVEID = $CVE.id
    $Summary = $CVE.summary
    $Output = "------------------------`nCVE No.: $Counter`n`nCVE ID: $CVEID`nSummary: $Summary`n------------------------`n"

    # If the user wants to save the CVE list to a .txt file, append the output to the file
    If ($SaveToFile -eq "yes" -or $SaveToFile -eq "y") {
        Add-Content -Path $FilePath -Value $Output
    }

    # Print the output to the console
    Write-Output $Output

    # Increment the counter
    $Counter++
}

#End message to the user
Write-Host "This is the end of the list`n`n"

# Pauses the PowerShell script so the user can read the results in the PowerShell window
Pause