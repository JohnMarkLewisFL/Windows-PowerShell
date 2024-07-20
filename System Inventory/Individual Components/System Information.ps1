# This script will retrieve some general information about the local computer and its current login session


# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# The ImportExcel module will allow the results to be exported properly into a .xlsx file type
Get-Module ImportExcel | Import-Module -Force

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

# Display the system information in the console
$SystemInfo | Format-List

# Prompt the user to save the system information
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$SaveFileDialog.Filter = "CSV files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx|Text files (*.txt)|*.txt"
$SaveFileDialog.FilterIndex = 1
$SaveFileDialog.RestoreDirectory = $true
$SaveFileDialog.FileName = "System Information - $Hostname"
$SaveFileDialog.Title = "Save As"

if ($SaveFileDialog.ShowDialog() -eq 'OK') {
    $FileType = $SaveFileDialog.FilterIndex
    $FilePath = $SaveFileDialog.FileName
    # Save the system information according to the selected file type
    switch ($FileType) {
        1 { $SystemInfo | Export-Csv -Path $FilePath -NoTypeInformation }
        2 { $SystemInfo | Export-Excel -Path $FilePath -AutoSize }
        3 { $SystemInfo | Out-File -FilePath $FilePath }
    }
}