# This script uses the PSWritePDF module to split a multi-page PDF file (one PDF file per page). Leading zeroes will be applied if there are ten or more pages in the original PDF file.

# This script requires the PSWritePDF module to be installed. The PSWritePDF module can be installed by running the following cmdlet as Administrator:
# Install-Module PSWritePDF

# Import the PSWritePDF module for PDF functionality
Import-Module PSWritePDF -Force

# Add the Windows Forms assembly for file open and file save dialog box functionality
Add-Type -AssemblyName System.Windows.Forms

# Introduction prompt and delay
Write-Host "Please select the PDF file to split. Each page will become its own PDF file.`n"
Write-Host "The split PDF files will have the original PDF file's name followed by the page number.`n`n"
Start-Sleep -Seconds 2

# Create an OpenFileDialog object
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "PDF files (*.pdf)|*.pdf"
$openFileDialog.Title = "Select a PDF File"


# Show the OpenFileDialog
if ($openFileDialog.ShowDialog() -eq "OK") {
    $pdfFilePath = $openFileDialog.FileName

    # Destination folder selection prompt and delay
    Write-Host "Please select the destination folder for the split PDF files.`n`n"
    Start-Sleep -Seconds 2

    # Create a FolderBrowserDialog object
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog

    # Show the FolderBrowserDialog
    if ($folderBrowserDialog.ShowDialog() -eq "OK") {
        $outputFolder = $folderBrowserDialog.SelectedPath

        # Get the base file name without extension
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($pdfFilePath)

        # Split the PDF file
        Split-PDF -FilePath $pdfFilePath -OutputFolder $outputFolder

        # Get the split PDF files
        $splitPdfFiles = Get-ChildItem -Path $outputFolder -Filter "OutputDocument*.pdf"

        # Determine the number of digits needed for the page number
        $numberOfDigits = [Math]::Floor([Math]::Log10($splitPdfFiles.Count)) + 1

        # Rename the split PDF files
        for ($i = 0; $i -lt $splitPdfFiles.Count; $i++) {
            # Calculate the page number with leading zeros
            $pageNumber = "{0:D$numberOfDigits}" -f ($i + 1)

            # Calculate the new file name
            $newFileName = "$baseFileName - Page $pageNumber.pdf"

            # Calculate the new file path
            $newFilePath = Join-Path -Path $outputFolder -ChildPath $newFileName

            # Rename the split PDF files
            Rename-Item -Path $splitPdfFiles[$i].FullName -NewName $newFilePath
        }
    }
}

# Final message showing the combined PDF file's save location
Write-Host "Your split PDF files have been saved at: $outputFolder`n`n"
Start-Sleep -Seconds 5