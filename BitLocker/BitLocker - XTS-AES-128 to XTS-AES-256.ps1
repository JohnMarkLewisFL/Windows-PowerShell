# This script will automatically check the C drive's current BitLocker status
# If 128-bit encryption is used, then it will automatically decrypt the drive 
# and prompt for a reboot to initiate BitLocker encryption with XTS-AES 256 (used space only)

# Elevate the script to run as Administrator automatically
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Introduction message
Write-Host "Checking BitLocker status on drive C:`n"
Start-Sleep -Seconds 2

# Retrieves the current BitLocker status for drive C:
$BitLockerStatus = Get-BitLockerVolume -MountPoint "C:"

# Function to force a reboot so BitLocker encryption can start
function Force-Reboot {
    Write-Host "`nRebooting in 5 seconds to start the encryption process."
    Start-Sleep -Seconds 5
    Restart-Computer -Force
}

# Case 1: Drive C: is not encrypted
if ($BitLockerStatus.EncryptionMethod -eq 'None') {
    Write-Host "`nDrive C: is not encrypted. Starting encryption with XTS-AES 256 (Used Space Only)."
    Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes256 -UsedSpaceOnly -RecoveryPasswordProtector
    Force-Reboot
}

# Case 2: Drive C: is encrypted with 128-bit encryption (XTS or legacy AES)
elseif ($BitLockerStatus.EncryptionMethod -in @('XtsAes128', 'Aes128')) {
    Write-Host "`nDrive C: is encrypted with 128-bit encryption. Starting decryption."
    Disable-BitLocker -MountPoint "C:"

    # Wait for the decryption process to complete
    do {
        Start-Sleep -Seconds 3
        $BitLockerStatus = Get-BitLockerVolume -MountPoint "C:"
        $Progress = 100 - $BitLockerStatus.EncryptionPercentage
        Write-Host "`nDecryption Progress: $progress% complete"
    } while ($BitLockerStatus.VolumeStatus -ne 'FullyDecrypted')

    Write-Host "`nDecryption complete. Starting encryption with XTS-AES 256 (Used Space Only)."
    Enable-BitLocker -MountPoint "C:" -EncryptionMethod XtsAes256 -UsedSpaceOnly -RecoveryPasswordProtector
    Force-Reboot
}

# Case 3: Drive C: is already encrypted with 256-bit encryption (XTS or legacy AES)
elseif ($BitLockerStatus.EncryptionMethod -in @('XtsAes256', 'Aes256')) {
    Write-Host "`nDrive C: is already encrypted with 256-bit encryption."
}

# Case 4: Unknown state
else {
    Write-Host "`nUnknown encryption state: $($bitlockerStatus.EncryptionMethod)"
}

Pause
