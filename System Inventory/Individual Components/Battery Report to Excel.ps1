# Load the Windows Forms assembly needed for the Save As file dialog box
Add-Type -AssemblyName System.Windows.Forms

# Imports the ImportExcel module needed for Excel functionality
Import-Module -Name ImportExcel

# Introduction message
Write-Host "This script will create a battery report using the powercfg command and save it as an Excel spreadsheet in a location of your choosing"`n
Start-Sleep -Seconds 2

# Generates the battery report as an XML file
& powercfg /batteryreport /XML /OUTPUT "batteryreport.xml"
Start-Sleep -Seconds 2

# Loads the XML file, parses the data, and formats it to be more Excel-friendly
[xml]$BatteryReport = Get-Content "batteryreport.xml"
$BatteryReportData = $BatteryReport.BatteryReport.Batteries | ForEach-Object {
    [PSCustomObject]@{
        DesignCapacity = $_.Battery.DesignCapacity
        FullChargeCapacity = $_.Battery.FullChargeCapacity
        BatteryHealth = [math]::floor([int64]$_.Battery.FullChargeCapacity/[int64]$_.Battery.DesignCapacity*100)
        CycleCount = $_.Battery.CycleCount
        ID = $_.Battery.id
    }
}

# Deletes the batteryreport.xml file
Remove-Item "batteryreport.xml"

# Creates the file save dialog box for the spreadsheet
Write-Host `n"You will now be prompted to save the battery report spreadsheet"
Start-Sleep -Seconds 3
$SaveAsBox = New-Object System.Windows.Forms.SaveFileDialog
$SaveAsBox.Filter = "Excel Files (*.xlsx)|*.xlsx"
$SaveAsBox.Title = "Save Battery Report as Excel Spreadsheet"

# Shows the save file dialog and exports the data to Excel if the user clicks OK
If ($SaveAsBox.ShowDialog() -eq "OK") {
    $BatteryReportData | Export-Excel -Path $SaveAsBox.FileName
}

# Completion message
Write-Host `n"Your battery report spreadsheet has been saved at: " $SaveAsBox.FileName
Start-Sleep -Seconds 5