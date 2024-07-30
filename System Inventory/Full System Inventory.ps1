# This script will retrieve some general information about the local Windows computer and its current login session
# It will also check a Windows host for a list of its Bluetooth adapters, Bluetooth devices, and network adapters


# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
# If the ImportExcel module is not installed, use the following cmdlet (may need to be run as administrator) to install the ImportExcel module:
# Install-Module ImportExcel
Get-Module ImportExcel | Import-Module -Force

# Write a message to the user
Write-Host "Please wait while the system information is gathered. The BitLocker portion will run as administrator.`n`n"

# Get the screen resolution
$Screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$ScreenWidth = $Screen.Width
$ScreenHeight = $Screen.Height

# The "Get-CimInstance xxx" commands retrieve the system information
$CIMComputerSystem = Get-CimInstance Win32_ComputerSystem
$CIMProcessor = Get-CimInstance CIM_Processor
$Win32_LogicalDisk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID = 'C:'"
$CDriveGB = [math]::Round($Win32_LogicalDisk.Size / 1GB)
$Win32_VideoController = Get-CimInstance Win32_VideoController
$Win32_BIOS = Get-CimInstance Win32_BIOS
$SerialNumber = $Win32_BIOS.SerialNumber
$BIOSVersion = $Win32_BIOS.SMBIOSBIOSVersion
$GPU = $Win32_VideoController.Name
$Hostname = $CIMComputerSystem.Name
$Domain = $CIMComputerSystem.Domain
$Username = $CIMComputerSystem.UserName
$Manufacturer = $CIMComputerSystem.Manufacturer
$Model = $CIMComputerSystem.Model
$CPU = $CIMProcessor.Name
$RAMBytes = $CIMComputerSystem.TotalPhysicalMemory
$RAMGBs = [math]::Round($RAMBytes / 1GB)

# Create an object to hold the system information and display each object in the following order
$SystemInfo = New-Object PSObject -Property ([ordered]@{
    Hostname = $Hostname
    "Domain/Workgroup" = $Domain
    "Current User" = $Username
    Manufacturer = $Manufacturer
    Model = $Model
    "Serial Number" = $SerialNumber
    "UEFI/BIOS Version" = $BIOSVersion
    CPU = $CPU
    GPU = $GPU
    RAM = "$RAMGBs GB"
    "C: Drive Capacity" = "$CDriveGB GB"
    "Screen Width" = $ScreenWidth
    "Screen Height" = $ScreenHeight
})

# Gets a list of the installed software
$InstalledSoftwareList = Get-CimInstance -Class Win32_Product | Select-Object Name,Version,Vendor,InstallDate,Description

# Gets all installed Bluetooth adapters and devices
$BluetoothList = Get-PnpDevice | Where-Object { $_.Class -eq "Bluetooth" } | Select-Object Name, Status

# Gets all installed network adapters
$NetworkAdapterList = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | 
                        ForEach-Object {
                            $ipv4 = $_.IPAddress | Where-Object { $_ -notmatch ":" }
                            $ipv6 = $_.IPAddress | Where-Object { $_ -match ":" }
                            $obj = New-Object PSObject -Property ([ordered]@{
                                Description = $_.Description
                                MACAddress = $_.MACAddress
                                IPAddressIPv4 = $ipv4
                                IPAddressIPv6 = $ipv6
                                IPSubnet = ($_.IPSubnet -join ", ")
                            })
                            $obj
                        }

# Gets all installed printers and their port names (software printers may show a strange port)
$PrintersList = Get-Printer | Select-Object Name, PortName

# Gets installed USB adapters and devices while filtering out most of the generic entries
$ExcludeUSBNames = "USB Composite Device", "USB Printing Support", "USB Mass Storage Device", "Generic SuperSpeed USB Hub", "USB Root Hub", "USB Root Hub (USB 3.0)", "Generic USB Hub"
$USBList = Get-PnpDevice | Where-Object { $_.Class -eq "USB" -and $ExcludeUSBNames -notcontains $_.Name } | Select-Object Name, Status

# Gets local and mapped drive information
$DriveList = Get-WmiObject -Class Win32_LogicalDisk
$DriveListFormatting = $DriveList | ForEach-Object {
    New-Object -TypeName PSObject -Property @{
        Username = $Username
        DriveLetter = $_.DeviceID
        Location = if ($_.DriveType -eq 4) { $_.ProviderName } else { $_.VolumeName }
    }
}

# Gather the SSID from the netsh command and split the unused text and whitespaces
$WirelessProfiles = netsh wlan show profiles | Select-String -Pattern "All User Profile" | %{ ($_ -split ":")[-1].Trim() };


# For each profile found with the netsh command, show the password in plaintext and display the SSID in columns
$WirelessProfiles | foreach {
	$ProfileData = netsh wlan show profiles name=$_ key="clear";
	$SSID = $ProfileData | Select-String -Pattern "SSID Name" | %{ ($_ -split ":")[-1].Trim().Trim('"') };
	$PSK = $ProfileData | Select-String -Pattern "Key Content" | %{ ($_ -split ":")[-1].Trim() };

    $WiFiList =	[PSCustomObject]@{
		SSID = $SSID;
		Password = $PSK
	    }
}

# Desktop folder path of the current user
$DesktopFolder = "$env:USERPROFILE\\Desktop"

# The BitLocker cmdlets must be run as administrator, so this will run only the BitLocker cmdlet as administrator
# This only gets the BitLocker information for drive C
$BitLockerCommand = "(Get-BitLockerVolume -MountPoint C).KeyProtector | Select-Object KeyProtectorId,KeyProtectorType,RecoveryPassword | Export-Csv -Path $DesktopFolder\\BitLockerTemp.csv -NoTypeInformation"
Start-Process powershell -Verb runAs -ArgumentList "-Command & {$BitLockerCommand}"
Start-Sleep -Seconds 2

# Saves the BitLocker cmdlet output as a .csv file so the rest of the script can use the output data
$BitLockerOutput = Import-Csv -Path $DesktopFolder\\BitLockerTemp.csv

# Retrieves the list of currently available power plans and format it
$PowerPlans = powercfg /list
$PowerPlanDetails = $PowerPlans -split "`n" | ForEach-Object { if ($_ -match 'Power Scheme GUID: (.+?)  \((.+)\)') { [PSCustomObject]@{'PowerScheme GUID'=$matches[1]; 'Plan Name'=$matches[2]} } }

# Generates the battery report as an XML file
& powercfg /batteryreport /XML /OUTPUT $DesktopFolder\\BatteryReport.xml
Start-Sleep -Seconds 2

# Loads the battery report XML file, parses the data, and formats it to be more Excel-friendly
[xml]$BatteryReport = Get-Content $DesktopFolder\\BatteryReport.xml
$BatteryReportData = $BatteryReport.BatteryReport.Batteries | ForEach-Object {
    [PSCustomObject]@{
        DesignCapacity = $_.Battery.DesignCapacity
        FullChargeCapacity = $_.Battery.FullChargeCapacity
        BatteryHealth = [math]::floor([int64]$_.Battery.FullChargeCapacity/[int64]$_.Battery.DesignCapacity*100)
        CycleCount = $_.Battery.CycleCount
        ID = $_.Battery.id
    }
}

# Prompt the user to save the results as an Excel spreadsheet (.xlsx) or .txt file
Write-Host "You will now be prompted to save the results as a spreadsheet or .txt file`n`n"
Start-Sleep -Seconds 3
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$SaveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Text file (*.txt)|*.txt"
$SaveFileDialog.FilterIndex = 1
$SaveFileDialog.RestoreDirectory = $true
$SaveFileDialog.FileName = "System Information - $Hostname"
$SaveFileDialog.Title = "Save Results Spreadsheet As"

if ($SaveFileDialog.ShowDialog() -eq 'OK') {
    $FileType = $SaveFileDialog.FilterIndex
    $FilePath = $SaveFileDialog.FileName

    # Save the system information according to the selected file type
    switch ($FileType) {
        1 { 
            $SystemInfo | Export-Excel -Path $FilePath -AutoSize -WorksheetName "System Info"
            $InstalledSoftwareList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Installed Software" -Append
            $DriveListFormatting | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Drive List" -Append
            $BitLockerOutput | Export-Excel -Path $FilePath -AutoSize -WorksheetName "BitLocker Recovery" -Append
            $PrintersList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Printers" -Append
            $WiFiList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Wi-Fi Networks" -Append
            $NetworkAdapterList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Network Adapters" -Append
            $BluetoothList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Bluetooth Devices" -Append
            $USBList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "USB Devices" -Append
            $PowerPlanDetails | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Power Plans" -Append
            $BatteryReportData | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Battery Health" -Append
        }
        2 { 
            $SystemInfo | Out-File -FilePath $FilePath
            $InstalledSoftwareList | Out-File -FilePath $FilePath -Append
            $DriveListFormatting | Out-File -FilePath $FilePath -Append
            $BitLockerOutput | Out-File -FilePath $FilePath -Append
            $BluetoothList | Out-File -FilePath $FilePath -Append
            $NetworkAdapterList | Out-File -FilePath $FilePath -Append
            $PrintersList | Out-File -FilePath $FilePath -Append
            $USBList | Out-File -FilePath $FilePath -Append
            $WiFiList | Out-File -FilePath $FilePath -Append
            $PowerPlanDetails | Out-File -FilePath $FilePath -Append
            $BatteryReportData | Out-File -FilePath $FilePath -Append
        }
    }
}

# Deletes the temporary files
Remove-Item -Path $DesktopFolder\\BitLockerTemp.csv
Remove-Item -Path $DesktopFolder\\BatteryReport.xml

# End message
Write-Host "Your results spreadsheet has been saved at: "$SaveFileDialog.FileName
Start-Sleep -Seconds 3
Write-Host "`n`nThis window will close shortly."
Start-Sleep -Seconds 3
