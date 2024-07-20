# This script is intended to perform a continuous ping to a destination, but show a notification window when the destination's state changes
# To stop this script, press Ctrl + C within the PowerShell window


# Adds the Windows Forms assembly to avoid potential errors when running the script
Add-Type -AssemblyName System.Windows.Forms

# Prompts the user for the remote IP address, hostname, or URL ($Destination)
$Destination = Read-Host 'Please enter the IP address, hostname (may need the domain suffix), or URL'

# The $PreviousState variable keeps track of the true/false nature of the success of the Test-NetConnection command
$PreviousState = $false


# The While loop and nested If Else statements run the $PingCommand while keeping track of the true/false for online and offline changes
While ($true) {
    # This command tests if the $Destination is online or offline and the -Quiet parameter is necessary for the true/false logic
    $PingCommand = Test-Connection -ComputerName $Destination -Quiet -Count 1

    # Get the current date and time
    $Timestamp = Get-Date

    # Display the output of the $PingCommand to the PowerShell window
    Write-Host $Destination 'ping succeeded:' $PingCommand $Timestamp

    If ($PingCommand -eq $true) {
        If (-not $PreviousState) {
            # Creates a popup message box
            [System.Windows.Forms.MessageBox]::Show($Destination + ' is currently online' , $Destination + ' Is Online' , 0, 48)
            $PreviousState = $true
            # Writes messages to the PowerShell window
            Write-Host 'You will not see another popup window until the destination goes offline'
            Write-Host 'Press Ctrl + C to stop the script if needed'
        }
    }
    Else {
        If ($PreviousState) {
            # Creates a popup message box
            [System.Windows.Forms.MessageBox]::Show($Destination + ' is currently offline' , $Destination + ' Is Offline' , 0, 48)
            $PreviousState = $false
            # Writes messaages to the PowerShell window
            Write-Host 'You will not see another popup window until the destination comes online'
            Write-Host 'Press Ctrl + C to stop the script if needed'
        }
    }
    # Pauses the $PingCommand in between runnings
    Start-Sleep -Seconds 5
}