# This script will test a connection to a user-specified port on a user-specified address


# Prompts the user for the remote IP address, hostname, or URL ($Destination)
$Destination = Read-Host 'Please enter the IP address, hostname (may need the the domain suffix), or URL'

# Prompts the user for the port number to test
$PortNumber = Read-Host 'Please enter the port number'
Test-NetConnection $Destination -Port $PortNumber

# Displays an end message once the script has finished running
Read-Host -Prompt 'The connection testing task has completed. Press the Enter key to close this window.'