# This script will retrieve some general information about the local Windows computer and its current login session
# It will also check a Windows host for a list of its Bluetooth adapters, Bluetooth devices, and network adapters


# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
# If the ImportExcel module is not installed, use the following cmdlet (may need to be run as administrator) to install the ImportExcel module:
# Install-Module ImportExcel
Get-Module ImportExcel | Import-Module -Force

# Write a message to the user
Write-Host "Please wait while the system information is gathered. The BitLocker portion will run as administrator.`n"

# Get the screen resolution
$Screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$ScreenWidth = $Screen.Width
$ScreenHeight = $Screen.Height

# The "Get-CimInstance xxx" cmdlets retrieve the system information
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
$InstalledSoftwareList = Get-CimInstance -Class Win32_Product | Select-Object Name,Version,Vendor,@{Name="Install Date";Expression={([datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null)).ToShortDateString()}},Description

# Gets all installed Bluetooth adapters and devices
$BluetoothList = Get-PnpDevice | Where-Object { $_.Class -eq "Bluetooth" } | Select-Object Name, Status

# Function to convert subnet masks to CIDR notation
Function Convert-SubnetMaskToCIDR {
    Param (
        [string]$SubnetMask
    )
    $Octets = $SubnetMask.Split('.')
    $CIDR = 0
    ForEach ($Octet In $Octets) {
        $CIDR += [Convert]::ToString([int]$octet, 2).Replace('0', '').Length
    }
    Return "/$CIDR"
}

# Function to convert CIDR to subnet masks
Function Convert-CIDRToSubnetMask {
    Param (
        [int]$CIDR
    )
    $Mask = [math]::Pow(2, 32) - [math]::Pow(2, 32 - $CIDR)
    $Bytes = [BitConverter]::GetBytes([UInt32]$mask)
    [IPAddress]::new($Bytes).IPAddressToString
}

# Gets all installed network adapters
$NetworkAdapterList = Get-NetIPConfiguration | ForEach-Object {
    $IPV4 = (Get-NetIPAddress -InterfaceIndex $_.InterfaceIndex -AddressFamily IPv4).IPAddress
    $IPV6 = (Get-NetIPAddress -InterfaceIndex $_.InterfaceIndex -AddressFamily IPv6).IPAddress
    $SubnetMaskV4 = Convert-CIDRToSubnetMask -CIDR (Get-NetIPAddress -InterfaceIndex $_.InterfaceIndex -AddressFamily IPv4).PrefixLength
    $SubnetCIDRV4 = Convert-SubnetMaskToCIDR -subnetMask $SubnetMaskV4
    $SubnetMaskV6 = (Get-NetIPAddress -InterfaceIndex $_.InterfaceIndex -AddressFamily IPv6).PrefixLength
    $SubnetCIDRV6 = "/$SubnetMaskV6"
    $MACAddress = (Get-NetAdapter -InterfaceIndex $_.InterfaceIndex).MacAddress
    $Obj = New-Object PSObject -Property ([ordered]@{
        Description = $_.InterfaceAlias
        "MAC Address" = $MACAddress
        "IPv4 Address" = $IPV4
        "IPv4 Subnet Mask" = $SubnetMaskV4
        "IPv4 Subnet CIDR" = $SubnetCIDRV4
        "IPv6 Address" = $IPV6
        "IPv6 Subnet Mask" = $SubnetMaskV6
        "IPv6 Subnet CIDR" = $SubnetCIDRV6
    })
    $Obj
}

# Gets all installed printers and their port names (software printers may show a strange port)
$PrintersList = Get-Printer | Select-Object Name,@{Name="Port Name";Expression={$_.PortName}}

# Gets installed USB adapters and devices while filtering out most of the generic entries
$ExcludeUSBNames = "USB Composite Device", "USB Printing Support", "USB Mass Storage Device", "Generic SuperSpeed USB Hub", "USB Root Hub", "USB Root Hub (USB 3.0)", "Generic USB Hub"
$USBList = Get-PnpDevice | Where-Object { $_.Class -eq "USB" -and $ExcludeUSBNames -notcontains $_.Name } | Select-Object Name, Status

# Gets local and mapped drive information
$DriveList = Get-WmiObject -Class Win32_LogicalDisk
$DriveListFormatting = $DriveList | ForEach-Object {
    New-Object -TypeName PSObject -Property @{
        Username = $Username
        "Drive Letter" = $_.DeviceID
        Location = If ($_.DriveType -eq 4) { $_.ProviderName } Else { $_.VolumeName }
    }
}

# Gather the SSID from the netsh command and split the unused text and whitespaces
$WirelessProfiles = netsh wlan show profiles | Select-String -Pattern "All User Profile" | %{ ($_ -split ":")[-1].Trim() };


# For each profile found with the netsh command, show the password in plaintext and display the SSID in columns
$WirelessProfiles | ForEach {
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
$PowerPlanDetails = $PowerPlans -split "`n" | ForEach-Object { if ($_ -match 'Power Scheme GUID: (.+?)  \((.+)\)') { [PSCustomObject]@{'Power Scheme GUID'=$matches[1]; 'Plan Name'=$matches[2]} } }

# Generates the battery report as an XML file
& powercfg /batteryreport /XML /OUTPUT $DesktopFolder\\BatteryReport.xml
Start-Sleep -Seconds 2

# Loads the battery report XML file, parses the data, and formats it to be more Excel-friendly
[xml]$BatteryReport = Get-Content "$DesktopFolder\BatteryReport.xml"
$BatteryReportData = $BatteryReport.BatteryReport.Batteries | ForEach-Object {
    [PSCustomObject]@{
        "Design Capacity" = $_.Battery.DesignCapacity
        "Full Charge Capacity" = $_.Battery.FullChargeCapacity
        "Battery Health" = [math]::floor([int64]$_.Battery.FullChargeCapacity/[int64]$_.Battery.DesignCapacity*100)
        "Cycle Count" = $_.Battery.CycleCount
        ID = $_.Battery.id
    }
}

# Retrieves all of the currently enabled firewall rules and formats the data
$FirewallRules = Get-NetFirewallRule | 
    Where-Object {$_.Enabled -eq "True"} | 
    Select-Object @{Name='Name';Expression={$_.Name}},
                  @{Name='Display Name';Expression={$_.DisplayName}},
                  @{Name='Enabled';Expression={$_.Enabled}},
                  @{Name='Profile';Expression={$_.Profile}},
                  @{Name='Direction';Expression={$_.Direction}},
                  @{Name='Action';Expression={$_.Action}},
                  @{Name='Edge Traversal Policy';Expression={$_.EdgeTraversalPolicy}},
                  @{Name='Description';Expression={$_.Description}},
                  @{Name='Display Group';Expression={$_.DisplayGroup}}

# Prompt the user to save the results as an Excel spreadsheet (.xlsx) or .txt file
Write-Host `n"You will now be prompted to save the results as a spreadsheet or .txt file"
Start-Sleep -Seconds 3
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$SaveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Text file (*.txt)|*.txt"
$SaveFileDialog.FilterIndex = 1
$SaveFileDialog.RestoreDirectory = $true
$SaveFileDialog.FileName = "System Information - $Hostname"
$SaveFileDialog.Title = "Save Results Spreadsheet As"

If ($SaveFileDialog.ShowDialog() -eq 'OK') {
    $FileType = $SaveFileDialog.FilterIndex
    $FilePath = $SaveFileDialog.FileName

    # Save the system information according to the selected file type
    Switch ($FileType) {
        1 { 
            $SystemInfo | Export-Excel -Path $FilePath -AutoSize -WorksheetName "System Info" -TableName "SystemInfo" -TableStyle Medium9
            $InstalledSoftwareList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Installed Software" -TableName "InstalledSoftware" -TableStyle Medium9 -Append
            $DriveListFormatting | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Drive List" -TableName "DriveList" -TableStyle Medium9 -Append
            $BitLockerOutput | Select-Object @{Name='Key Protector Id';Expression={$_.KeyProtectorId}}, @{Name='Key Protector Type';Expression={$_.KeyProtectorType}}, @{Name='Recovery Password';Expression={$_.RecoveryPassword}} | Export-Excel -Path $FilePath -AutoSize -WorksheetName "BitLocker Recovery" -TableName "BitLockerRecovery" -TableStyle Medium9 -Append
            $PrintersList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Printers" -TableName "Printers" -TableStyle Medium9 -Append
            $WiFiList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Wi-Fi Credentials" -TableName "WiFiCredentials" -TableStyle Medium9 -Append
            $NetworkAdapterList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Network Adapters" -TableName "NetworkAdapters" -TableStyle Medium9 -Append
            $FirewallRules | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Firewall Rules" -TableName "FirewallRules" -TableStyle Medium9 -Append
            $BluetoothList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Bluetooth Devices" -TableName "BluetoothDevices" -TableStyle Medium9 -Append
            $USBList | Export-Excel -Path $FilePath -AutoSize -WorksheetName "USB Devices" -TableName "USBDevices" -TableStyle Medium9 -Append
            $PowerPlanDetails | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Power Plans" -TableName "PowerPlans" -TableStyle Medium9 -Append
            $BatteryReportData | Export-Excel -Path $FilePath -AutoSize -WorksheetName "Battery Health" -TableName "BatteryHealth" -TableStyle Medium9 -Append
        }
        2 { 
            $SystemInfo | Out-File -FilePath $FilePath
            $InstalledSoftwareList | Out-File -FilePath $FilePath -Append
            $DriveListFormatting | Out-File -FilePath $FilePath -Append
            $BitLockerOutput | Out-File -FilePath $FilePath -Append
            $BluetoothList | Out-File -FilePath $FilePath -Append
            $NetworkAdapterList | Out-File -FilePath $FilePath -Append
            $FirewallRules | Out-File -FilePath $FilePath -Append
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
Write-Host `n"Your results file has been saved at: $FilePath"
Start-Sleep -Seconds 3
Write-Host `n"This window will close shortly."
Start-Sleep -Seconds 5
