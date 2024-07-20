# This script will ping a remote host every 2 seconds for 2147483647 times (the int max amount)
# When running this script, the user will be prompted to enter either the hostname or IP address of the remote machine to ping


# Automatically runs the maximum number of times possible for int values
$PingCount = 2147483647
$Results = @()

# Prompts the user for the remote IP address, hostname, or URL ($Destination)
$Destination = Read-Host 'Please enter the IP address, hostname (may need the domain suffix), or URL'


# Pings the remote host and adds timestamp, IPv4 address, and response time
For ($i = 1; $i -le [int]$PingCount; $i++) {
    Try {
        $Result = Test-Connection -ComputerName $Destination -Count 1 -ErrorAction Stop | 
                  Select-Object @{n='Timestamp';e={Get-Date}},Address,IPv4Address,IPv6Address,ResponseTime,
                                @{n='Status';e={'Success'}}
    }
    Catch {
        $Result = New-Object PSObject -Property @{
            Timestamp = Get-Date
            Address = $Destination
            IPv4Address = $null
            IPv6Address = $null
            ResponseTime = $null
            Status = 'Fail'
        }
    }
    $Results += $Result
    $Result | Format-Table
    Start-Sleep -Seconds 2
}


# Displays an end message once the script has finished running
Read-Host -Prompt 'The ping task has completed. Press the Enter key to close this window.'