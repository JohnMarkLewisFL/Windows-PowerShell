# This script can help semi-automate BitLocker encryption and ensure XTS-AES 256 is the encryption method used
# This script tries to focus on PowerShell cmdlets as opposed to the older (Command Prompt-friendly) manage-bde command
# If you remove the "run as administrator" portion of this script, then the script will need to be run from an Administrator Windows PowerShell window

# Elevates the script to run as Administrator automatically
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" `"$args`"" -Verb RunAs
    exit
}

# Prompts the user to enter a drive letter and validates the input with a do while loop
do {
    $Drive = Read-Host "`nPlease enter a valid drive letter (A-Z or a-z)"
    if ($Drive -match "^[a-zA-Z]$") {
        $Drive = $Drive.ToUpper()
        $Valid = $true
    } else {
        Write-Host "`nInvalid input. Please enter a single letter from A to Z."
        $Valid = $false
    }
} while (-not $Valid)

# Function that displays the "main menu" of the script
function Show-Menu {
    Clear-Host
    Write-Host "==============================="
    Write-Host " BitLocker Toolbox for ${Drive}:"
    Write-Host "==============================="
    Write-Host "`n 1. Check BitLocker Status"
    Write-Host " 2. Decrypt Drive"
    Write-Host " 3. Encrypt with XTS-AES 256 - Used Space Only (Fast, recommended for fixed drives)"
    Write-Host " 4. Encrypt with XTS-AES 256 - Full Disk (Slow, recommended for fixed drives)"
    Write-Host " 5. Encrypt with XTS-AES 128 - Used Space Only (Fast, recommended for fixed drives)"
    Write-Host " 6. Encrypt with XTS-AES 128 - Full Disk (Slow, recommended for fixed drives)"
    Write-Host " 7. Encrypt with AES 256 - Used Space Only (Fast, recommended for removable drives)"
    Write-Host " 8. Encrypt with AES 256 - Full Disk (Slow, recommended for removable drives)"
    Write-Host " 9. Encrypt with AES 128 - Used Space Only (Fast, recommended for removable drives)"
    Write-Host "10. Encrypt with AES 128 - Full Disk (Slow, recommended for removable drives)"
    Write-Host "11. Exit"
    Write-Host "`n==============================="
}

# Function to retrieve the drive's current BitLocker status
function Get-DriveStatus {
    return Get-BitLockerVolume -MountPoint $Drive
}

# Function to check the current decryption or encryption percentage using the Get-DriveStatus PowerShell cmdlet
function Check-BitLockerStatus {
    $Status = Get-DriveStatus
    Write-Host "`n--- BitLocker Status for Drive $Drive ---"
    Write-Host "`nProtection Status : $($Status.ProtectionStatus)"
    Write-Host "Volume Status     : $($Status.VolumeStatus)"
    Write-Host "Encryption Method : $($Status.EncryptionMethod)"
    Write-Host "Encryption %      : $($Status.EncryptionPercentage)%"
    Write-Host "Key Protector Type: $($Status.KeyProtector.KeyProtectorType)"
    Write-Host "Key Protector ID  : $($Status.KeyProtector.KeyProtectorId)"
    Write-Host "Recovery Password : $($Status.KeyProtector.RecoveryPassword)"
    Write-Host "`n------------------------------------------`n"
    Pause
}

# Function that decrypts the drive using the Disable-BitLocker PowerShell cmdlet
function Decrypt-Drive {
    $Status = Get-DriveStatus
    if ($Status.ProtectionStatus -eq 'Off' -and $Status.VolumeStatus -eq 'FullyDecrypted') {
        Write-Host "`nDrive is currently decrypted"
    } else {
        Write-Host "`nDecrypting drive $Drive"
        Disable-BitLocker -MountPoint $Drive
        Write-Host "`nDecryption status updates will appear shortly`n"
        do {
            Start-Sleep -Seconds 5
            $Status = Get-DriveStatus
            $DecryptionPercentage = 100 - $Status.EncryptionPercentage
            Write-Host "Decryption Progress: $DecryptionPercentage% complete"
        } while ($Status.ProtectionStatus -ne 'Off' -or $Status.VolumeStatus -ne 'FullyDecrypted')
        Write-Host "`nDecryption complete."
    }
    Pause
}

# Function to prompt the user to automatically reboot now or reboot manually later
function RebootPrompt {
    do {
        $RebootChoice = Read-Host "`nA reboot is required to proceed with encryption. Reboot now? (Y/N)"
    } while ($RebootChoice -notmatch '^[YyNn]$')
    if ($RebootChoice -match '^[Yy]$') {
        Write-Host "`nRebooting in 3 seconds"
        Start-Sleep -Seconds 3
        Restart-Computer
    } else {
        Write-Host "`nYou have chosen not to reboot. Please note that encryption will not commence until the next reboot."
    }
}

# Function to encrypt the drive with a specified method and mode
function Encrypt-Drive($Method, $Mode) {
    $Status = Get-DriveStatus
    if ($Status.ProtectionStatus -eq 'On') {
        Write-Host "`nDrive $Drive is already encrypted. Please decrypt the drive before continuing."
    } else {
        if ($Mode -eq "used") {
            Write-Host "`nEncrypting used space only with $Method on $Drive"
            Enable-BitLocker -MountPoint $Drive -EncryptionMethod $Method -RecoveryPasswordProtector -UsedSpaceOnly
        } elseif ($Mode -eq "full") {
            Write-Host "`nEncrypting the full drive with $Method on $Drive"
            Enable-BitLocker -MountPoint $Drive -EncryptionMethod $Method -RecoveryPasswordProtector
        }
        Start-Sleep -Seconds 3
        RebootPrompt
    }
    Pause
}

# Do while loop for the main menu
do {
    Show-Menu
    $Choice = Read-Host "Select an option (1-11)"

    switch ($Choice) {
        '1' { Check-BitLockerStatus }
        '2' { Decrypt-Drive }
        '3' { Encrypt-Drive "XtsAes256" "used" }
        '4' { Encrypt-Drive "XtsAes256" "full" }
        '5' { Encrypt-Drive "XtsAes128" "used" }
        '6' { Encrypt-Drive "XtsAes128" "full" }
        '7' { Encrypt-Drive "Aes256" "used" }
        '8' { Encrypt-Drive "Aes256" "full" }
        '9' { Encrypt-Drive "Aes128" "used" }
        '10' { Encrypt-Drive "Aes128" "full" }
        '11' { Write-Host "Exiting the script"; exit }
        default {
            Write-Warning "`nInvalid input. Please select a number between 1 and 11."
            Pause
        }
    }
} while ($true)
