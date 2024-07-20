# This PowerShell script will automatically enable System Restore on the C drive, set the disk usage percentage to a user-specified value, and create a restore point

# Run as administrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$args`"" -Verb RunAs; exit }

# Introduction message
Write-Host "This script will enable System Restore on the C drive, set the usage to a percentage you specify, and create a restore point."
Write-Host `n`n"This script is running as administrator."`n`n
Start-Sleep -Seconds 3

# Get the disk space percentage from the user and validate the input so it is only a number from 1 to 100
Do {
    $MaxSizePercentage = Read-Host "Please enter the max disk space percentage for your system restore points [1-100]"
    $MaxSizePercentage = $MaxSizePercentage.Replace('%','')

    # Validate the input
    [int]$number = 0
    If (![int]::TryParse($MaxSizePercentage, [ref]$number) -or $number -lt 1 -or $number -gt 100) {
        Write-Host "Invalid input. Please enter a number between 1 and 100."
    }
} While ($number -lt 1 -or $number -gt 100)

# Formatting
Write-Host `n`n

# Enable System Restore on the C drive
Enable-ComputerRestore -Drive "C:\"

# Set disk space usage to the input percentage
vssadmin Resize ShadowStorage /For=C: /On=C: /MaxSize=$MaxSizePercentage%

# Create a restore point
Checkpoint-Computer -Description "Created By PowerShell Script" -RestorePointType MODIFY_SETTINGS

# Completion message
Write-Host "`n`nThe script has completed. This window will close shortly."
Start-Sleep -Seconds 5