# This script uses the PSWritePDF module to combine multiple PDF files into a single PDF file

# This script requires the PSWritePDF module to be installed. The PSWritePDF module can be installed by running the following cmdlet as Administrator:
# Install-Module PSWritePDF

# Import the PSWritePDF module for PDF functionality
Import-Module PSWritePDF -Force

# Add the Windows Forms assembly for file open and file save dialog box functionality
Add-Type -AssemblyName System.Windows.Forms

# Introduction prompt and 5 second delay
Write-Host "Please select the PDF files you need to combine. You will be prompted to select each file individually. You will need to select them in the proper order as you will not be able to rearrange the selected PDF files.`n`n"
Start-Sleep -Seconds 5

# Creates an ArrayList to hold the selected PDF files
$PDFFiles = New-Object System.Collections.ArrayList

# Loop to add PDF all files
While ($True) {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "PDF files (*.pdf)|*.pdf"
    $OpenFileDialog.Multiselect = $False
    $OpenFileDialog.Title = "Select a PDF file"

    If ($OpenFileDialog.ShowDialog() -eq "OK") {
        $PDFFiles.Add($OpenFileDialog.FileName) | Out-Null
        Write-Host "Here is the current order of the $($PDFFiles.Count) PDF files you selected:`n"
        $PDFFiles | ForEach-Object { Write-Host $_ }
        Write-Host "`n`n"

        Do {
            $Result = Read-Host "Do you want to keep the current order or remove the most recently added PDF file? (Keep/Remove)"
            Write-Host "`n`n"
            $Result = $Result.ToLower()
        } Until ($Result -eq "keep" -or $Result -eq "k" -or $Result -eq "remove" -or $Result -eq "r")

        If ($Result -eq "remove" -or $Result -eq "r") {
            $PDFFiles.RemoveAt($PDFFiles.Count - 1)
            Write-Host "The most recently added PDF file has been removed.`n`n"
            Write-Host "Here is the current order of the $($PDFFiles.Count) PDF files you selected:`n"
        $PDFFiles | ForEach-Object { Write-Host $_ }
        Write-host "`n`n"
        }
    }

    Do {
        $Result = Read-Host "Do you want to add another PDF file? (Yes/No)"
        Write-Host "`n`n"
        $Result = $Result.ToLower()
    } Until ($Result -eq "yes" -or $Result -eq "y" -or $Result -eq "no" -or $Result -eq "n")

    If ($Result -eq "no" -or $Result -eq "n") {
        Break
    }
}

# Check if any PDF files were selected
If ($PDFFiles.Count -eq 0) {
    Write-Host "No PDF files were selected.`n`n"
    Return
}

# Prompt for the save location of the combined PDF file
Write-Host "Your PDF file selection is complete. You will now be prompted to select a save location and name the combined PDF file.`n`n"
Start-Sleep -Seconds 5

$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.Filter = "PDF files (*.pdf)|*.pdf"

If ($SaveFileDialog.ShowDialog() -eq "OK") {
    $OutputFile = $SaveFileDialog.FileName
} Else {
    Write-Host "No output file was selected.`n`n"
    Return
}

# Combines the PDF files
Merge-PDF -InputFile $PDFFiles -OutputFile $OutputFile

# Final message showing the combined PDF file's save location
Write-Host "Your combined PDF file has been saved at: $OutputFile`n`n"
Start-Sleep -Seconds 5