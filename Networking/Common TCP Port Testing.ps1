# This script will test if a list of common TCP ports (not UDP) are open on a single host or a user-provided list of hosts


# Imports the Import-Excel module necessary for the Excel functionality
Import-Module ImportExcel

# Loads the Windows Forms assembly needed for the file dialog boxes
Add-Type -AssemblyName System.Windows.Forms

# Introduction message
Write-Host "This script will test a list of common TCP ports on a remote host or a list of remote hosts"
Write-Host "Due to the nature of the Test-NetConnection cmdlet, this will only test TCP and not UDP"
Start-Sleep -Seconds 2

# Function to prompt the user for the target host input and choice validation
Function Get-TargetHosts {
    While ($True) {
        Write-Host `n"Choose option 1 to manually enter a single target host to test"
        Write-Host "Choose option 2 to select a list of target hosts from a file (.txt, .csv, or .xlsx)"
        Write-Host "CSV files and Excel spreadsheets will need to have a colunm named TargetHost to process the target host list"
        Start-Sleep -Seconds 2
        $UserChoice = Read-Host `n"Enter your choice"
        If ($UserChoice -eq '1') {
            Return @(Read-Host "Enter the target hostname or IP address")
        } ElseIf ($UserChoice -eq '2') {
            $ListDialogBox = New-Object System.Windows.Forms.OpenFileDialog
            $ListDialogBox.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx"
            $ListDialogBox.Title = "Select a list of target hosts"
            $ListDialogBox.ShowDialog() | Out-Null
            $OpenFilePath = $ListDialogBox.FileName
            If ($OpenFilePath.EndsWith(".txt")) {
                Return Get-Content -Path $OpenFilePath
            } ElseIf ($OpenFilePath.EndsWith(".csv")) {
                Return Import-Csv -Path $OpenFilePath | Select-Object -ExpandProperty TargetHost
            } ElseIf ($OpenFilePath.EndsWith(".xlsx")) {
                Return Import-Excel -Path $OpenFilePath | Select-Object -ExpandProperty TargetHost
            }
        } Else {
            Write-Host `n"Invalid choice. Please try again."
        }
    }
}

# Prompts the user for the target host or target host list
$TargetHosts = Get-TargetHosts

# Prompts user to select a save location for the results spreadsheet
Write-Host `n"Please select the save location for the results spreadsheet"`n
Start-Sleep -Seconds 3
$ResultsDialogBox = New-Object System.Windows.Forms.SaveFileDialog
$ResultsDialogBox.Filter = "Excel Files (*.xlsx)|*.xlsx"
$ResultsDialogBox.FileName = "Port Testing Results.xlsx"
$ResultsDialogBox.Title = "Save the results spreadsheet"
$ResultsDialogBox.ShowDialog() | Out-Null
$ResultsFilePath = $ResultsDialogBox.FileName

# List of common TCP ports
$Ports = @(
    @{ Port = 20; Description = "FTP Data Transfer" },
    @{ Port = 21; Description = "FTP Command Control" },
    @{ Port = 22; Description = "SSH" },
    @{ Port = 23; Description = "Telnet" },
    @{ Port = 25; Description = "SMTP" },
    @{ Port = 53; Description = "DNS" },
    @{ Port = 67; Description = "DHCP Server" },
    @{ Port = 68; Description = "DHCP Client" },
    @{ Port = 69; Description = "TFTP" },
    @{ Port = 80; Description = "HTTP" },
    @{ Port = 110; Description = "POP3" },
    @{ Port = 119; Description = "NNTP" },
    @{ Port = 123; Description = "NTP" },
    @{ Port = 143; Description = "IMAP" },
    @{ Port = 161; Description = "SNMP" },
    @{ Port = 194; Description = "IRC" },
    @{ Port = 443; Description = "HTTPS" },
    @{ Port = 465; Description = "SMTPS" },
    @{ Port = 514; Description = "Syslog" },
    @{ Port = 587; Description = "SMTP (Mail Submission)" },
    @{ Port = 636; Description = "LDAPS" },
    @{ Port = 691; Description = "Microsoft Exchange" },
    @{ Port = 860; Description = "iSCSI" },
    @{ Port = 873; Description = "rsync" },
    @{ Port = 902; Description = "VMware ESXi" },
    @{ Port = 989; Description = "FTPS" },
    @{ Port = 990; Description = "FTPS" },
    @{ Port = 993; Description = "IMAPS" },
    @{ Port = 995; Description = "POP3S" },
    @{ Port = 1433; Description = "Microsoft SQL Server" },
    @{ Port = 1521; Description = "Oracle Database" },
    @{ Port = 3306; Description = "MySQL Database Server" },
    @{ Port = 3389; Description = "RDP" },
    @{ Port = 5432; Description = "PostgreSQL Database Server" },
    @{ Port = 5900; Description = "VNC" },
    @{ Port = 8006; Description = "HTTPS Alternate" }
    @{ Port = 8080; Description = "HTTP Alternate" },
    @{ Port = 8443; Description = "HTTPS Alternate" },
    @{ Port = 27017; Description = "MongoDB Database Server" }
)

# Creates a runspace pool to help with threading and faster execution
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount)
$RunspacePool.Open()

ForEach ($TargetHost in $TargetHosts) {
    Write-Host "Starting port testing on $TargetHost"
    # Array to store the runspaces
    $Runspaces = @()

    # Runs Test-NetConnection in parallel for each port in the list
    ForEach ($PortInfo in $Ports) {
        $Runspace = [powershell]::Create().AddScript({
            param ($Port, $PortDescription, $TargetHost)
            $Result = Test-NetConnection -ComputerName $TargetHost -Port $Port
            [PSCustomObject]@{
                'Target Host' = $TargetHost
                Port = $Port
                Description = $PortDescription
                Result = If ($Result.TcpTestSucceeded) { "Success" } Else { "Fail" }
            }
        }).AddArgument($PortInfo.Port).AddArgument($PortInfo.Description).AddArgument($TargetHost)
        $Runspace.RunspacePool = $RunspacePool
        $Runspaces += [PSCustomObject]@{ Pipe = $Runspace; Status = $Runspace.BeginInvoke() }
    }

    # Initializes the progress bar (otherwise the PowerShell console window looks frozen)
    $Progress = 0
    $TotalPorts = $Ports.Count

    # Waits for all runspaces to complete and collects the results
    $Results = ForEach ($Runspace in $Runspaces) {
        $Runspace.Pipe.EndInvoke($Runspace.Status)
        $Runspace.Pipe.Output | Where-Object { $_ -ne $Null }
        $Progress++
        Write-Progress -Activity "Testing Ports on $TargetHost" -Status "Processing port $Progress of $TotalPorts" -PercentComplete (($Progress / $TotalPorts) * 100)
    }

    # Formats the table names in the spreadsheet because Excel is picky with table names
    $FormattedTargetHost = $TargetHost -replace '[^a-zA-Z0-9]', '_'
    $TableName = "Results_$FormattedTargetHost"

    # Exports the port testing results to Excel file
    $WorksheetName = "$TargetHost Results"
    $Results | Export-Excel -Path $ResultsFilePath -WorksheetName $WorksheetName -TableName $TableName -TableStyle Medium9 -AutoSize
}

# Closes the runspace pool
$RunspacePool.Close()
$RunspacePool.Dispose()

# Ending message to the user
Write-Host `n"TCP port testing has completed"
Write-Host "The port testing results spreadsheet was saved to $ResultsFilePath"
Start-Sleep -Seconds 3
Write-Host `n"This window will close shortly"
Start-Sleep -Seconds 5