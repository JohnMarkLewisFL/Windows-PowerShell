# This script is designed to continuously ping a destination (a number of times defined by the user) and log the results to a file


# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
# If there are error messages regarding the 'Export-Excel' cmdlet, then run the following line in a separate PowerShell window to install the ImportExcel module manually:
# Install-Module -Name ImportExcel
Get-Module ImportExcel | Import-Module -Force

# Introduction message
Write-Host "This script will continuously ping a destination for a number of times you specify and log the results"
Start-Sleep -Seconds 2
Write-Host `n"You will now be prompted to save the log file and select its file type"`n
Start-Sleep -Seconds 3

# Creates the file save dialog box
Add-Type -AssemblyName System.Windows.Forms
$DialogBox = New-Object System.Windows.Forms.SaveFileDialog
$DialogBox.Filter = "Excel spreadsheet (*.xlsx)|*.xlsx|CSV file (*.csv)|*.csv|Text file (*.txt)|*.txt"
$DialogBox.Title = "Save Log As"

# Get the current date and time, format it as a string, and include it in the file name
$CurrentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$DialogBox.FileName = "$CurrentDateTime - Ping Results"
$DialogBox.ShowDialog() | Out-Null
$FilePath = $DialogBox.FileName


# Prompt the user for the destination to ping
$Destination = Read-Host -Prompt 'Enter the IP address, hostname (may require the domain name), or URL'


# Prompt the user for the number of pings and validate the input
Do {
    $PingCount = Read-Host -Prompt `n"Enter the number of times to ping the destination"
    If (![int]::TryParse($PingCount, [ref]$null) -or [int]$PingCount -lt 1 -or [int]$PingCount -gt 2147483647) {
        Write-Host `n"Input error. Please enter a number between 1 and 2147483647."
    }
} While (![int]::TryParse($PingCount, [ref]$null) -or [int]$PingCount -lt 1 -or [int]$PingCount -gt 2147483647)


# Initialize an array to hold the results
$Results = @()

# Perform the pings and save the results to the array
For ($i = 1; $i -le [int]$PingCount; $i++) {
    Try {
        # Get the MAC address using arp command
        $ArpResult = arp -a $Destination
        $MacAddress = ($ArpResult -split "`n" | Where-Object { $_ -match $Destination }) -replace ".+\s([\dA-Fa-f-]+)\s+.+", '$1'

        $Result = Test-Connection -ComputerName $Destination -Count 1 -ErrorAction Stop | 
                  Select-Object @{n='Timestamp';e={Get-Date}},Address,@{n='MACAddress';e={$MacAddress}},IPv4Address,IPv6Address,ResponseTime,
                                @{n='Status';e={'Success'}}
    }
    Catch {
        $Result = New-Object PSObject -Property @{
            Timestamp = Get-Date
            Address = $Destination
            MACAddress = $null
            IPv4Address = $null
            IPv6Address = $null
            ResponseTime = $null
            Status = 'Fail'
        }
    }
    $Results += $Result

    # Shows a progress bar, but this does not show the actual Test-Connection command output as it runs, only the number of pings completed out of the total number of pings
    Write-Progress -Activity "Pinging $Destination" -Status "$i of $PingCount pings completed" -PercentComplete ($i / $PingCount * 100)

    # Add a 2 second delay between pings, otherwise all pings are executed at the exact same second
    If ($i -lt [int]$PingCount) {
        Start-Sleep -Seconds 2
    }
}

# Save the results to the selected file type
If ($FilePath -match ".xlsx$") {
    $Results | Export-Excel -Path $FilePath -TableName "PingResults" -TableStyle Medium9
} Elseif ($FilePath -match ".csv$") {
    $Results | Export-Csv -Path $FilePath -NoTypeInformation
} Elseif ($FilePath -match ".txt$") {
    $Results | Out-File -FilePath $FilePath
}

# Conclusion message
Write-Host `n"The continuous ping operation has completed"
Write-Host `n`n"The results have been saved at the following location: $FilePath"
Start-Sleep -Seconds 5
