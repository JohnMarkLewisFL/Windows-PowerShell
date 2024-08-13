# This script will automatically retrieve the EXIF data from pictures and log it to an Excel spreadsheet
# There is a ton of extra code due to some of the options as well as GPS data conversions. It could probably be simplified in the right hands.


# Loads the Windows Forms assembly required for the file dialog boxes
Add-Type -AssemblyName System.Windows.Forms

# Introduction message
Write-Host "This PowerShell script will retrieve the EXIF data from pictures and log the results to an Excel spreadsheet"
Start-Sleep -Seconds 2

# Prompt the user to select single file or folder
Do {
    Write-Host `n"Enter 1 to process a single file"
    Write-Host "Enter 2 to process a folder"
    $FileOrFolderChoice = Read-Host `n"Please enter your choice for single file or folder"
} While ($FileOrFolderChoice -ne '1' -and $FileOrFolderChoice -ne '2')

Start-Sleep -Seconds 2

# Prompt the user to select recursion or no recursion if folder is selected
If ($FileOrFolderChoice -eq '2') {
    Do {
        Write-Host `n"Enter 1 to use recursion (process subfolders)"
        Write-Host "Enter 2 to not use recursion (do not process subfolders)"
        $RecursionChoice = Read-Host `n"Please enter your choice for recursion"
    } While ($RecursionChoice -ne '1' -and $RecursionChoice -ne '2')
}

# File dialog box for selecting a file or folder
If ($FileOrFolderChoice -eq '1') {
    Write-Host `n"You will now be prompted to select a file to process"
    Start-Sleep -Seconds 3
    $FileSelection = New-Object System.Windows.Forms.OpenFileDialog
    $FileSelection.Filter = "Image Files|*.jpg;*.jpeg;*.raw;*.tiff"
    $FileSelection.Title = "Select a single picture to process"
    $Null = $FileSelection.ShowDialog()
    $SelectedPath = $FileSelection.FileName
} Else {
    Write-Host `n"You will now be prompted to select a folder to process"
    Start-Sleep -Seconds 3
    $FolderSelection = New-Object System.Windows.Forms.FolderBrowserDialog
    $Null = $FolderSelection.ShowDialog()
    $SelectedPath = $FolderSelection.SelectedPath
}

# Function to convert GPS coordinates from EXIF format to decimal, otherwise garbage characters or a blank will be displayed
Function Convert-GPSCoordinate {
    Param (
        [byte[]]$Coordinate,
        [string]$Ref
    )

    $Degrees = [BitConverter]::ToUInt32($Coordinate[0..3], 0) / [BitConverter]::ToUInt32($Coordinate[4..7], 0)
    $Minutes = [BitConverter]::ToUInt32($Coordinate[8..11], 0) / [BitConverter]::ToUInt32($Coordinate[12..15], 0)
    $Seconds = [BitConverter]::ToUInt32($Coordinate[16..19], 0) / [BitConverter]::ToUInt32($Coordinate[20..23], 0)

    $DecimalCoordinate = $Degrees + ($Minutes / 60) + ($Seconds / 3600)

    If ($Ref -eq "S" -or $Ref -eq "W") {
        $DecimalCoordinate = -$DecimalCoordinate
    }

    Return $DecimalCoordinate
}

# Function to process a single file if this option is chosen by the user
Function Process-File {
    Param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )

    # Creates a Shell.Application COM object
    $Shell = New-Object -ComObject Shell.Application
    # Gets the folder containing the file
    $Folder = $Shell.Namespace((Get-Item $FilePath).DirectoryName)
    # Gets the file object
    $File = $Folder.ParseName((Get-Item $FilePath).Name)

    # Checks if the file object is not null
    If ($File) {
        # Retrieve and store all the details
        $FileData = @{}
        For ($i = 0; $i -lt 266; $i++) {
            $Detail = $Folder.GetDetailsOf($File, $i)
            If ($Detail) {
                $Property = $Folder.GetDetailsOf($Folder.Items, $i)
                # Excludes the following fields since they are redundant, but they can easily be added back if needed
                If ($Property -notin @("Space used", "Owner", "Folder", "Date accessed", "Link status", "Shared", "Computer", "Total size", "Folder name", "Space free", "Type", "Rating", "Kind", "Perceived type", "Item type", "Name")) {
                    $FileData[$Property] = $Detail -replace ';', ''
                }
            }
        }

        # Process GPS data separately using System.Drawing
        $Image = [System.Drawing.Image]::FromFile($FilePath)
        Try {
            $GPSLatitude = $Image.GetPropertyItem(2).Value
            $GPSLatitudeRef = $Image.GetPropertyItem(1).Value
            $GPSLongitude = $Image.GetPropertyItem(4).Value
            $GPSLongitudeRef = $Image.GetPropertyItem(3).Value
            $GPSAltitude = $Image.GetPropertyItem(6).Value

            $FileData["Latitude"] = Convert-GPSCoordinate -Coordinate $GPSLatitude -Ref ([System.Text.Encoding]::ASCII.GetString($GPSLatitudeRef))
            $FileData["Longitude"] = Convert-GPSCoordinate -Coordinate $GPSLongitude -Ref ([System.Text.Encoding]::ASCII.GetString($GPSLongitudeRef))
            $FileData["Altitude"] = [BitConverter]::ToUInt32($GPSAltitude, 0) / [BitConverter]::ToUInt32($GPSAltitude[4..7], 0)
        } Catch {
            Write-Output `n"Error: Unable to retrieve GPS data."
        }

        $Null = $EXIFData.Add([PSCustomObject]$FileData)
    } Else {
        Write-Output `n"Error: Unable to retrieve file details."
    }
}

# Function to process folders recursively
Function Process-Folder {
    Param (
        [Parameter(Mandatory=$True)]
        [string]$FolderPath
    )

    # Creates a Shell.Application COM object
    $Shell = New-Object -ComObject Shell.Application
    # Gets the folder containing photos
    $Folder = $Shell.Namespace($FolderPath)

    # Checks if the folder object is not null
    If ($Folder) {
        # Loops through all items in the folder
        ForEach ($Item in $Folder.Items()) {
            # If the item is a folder, call the function recursively
            If ($Item.IsFolder) {
                $null = Process-Folder -FolderPath $Item.Path
            } Else {
                # Filters file types
                If ($Item.Name -match '\.(jpg|jpeg|raw|tiff)$') {
                    # Retrieves and stores all the details
                    $FileData = @{}
                    For ($i = 0; $i -lt 266; $i++) {
                        $Detail = $Folder.GetDetailsOf($Item, $i)
                        If ($Detail) {
                            $Property = $Folder.GetDetailsOf($Folder.Items, $i)
                            # Excludes the following fields since they are redundant, but they can easily be added back if needed
                            If ($Property -notin @("Space used", "Owner", "Folder", "Date accessed", "Link status", "Shared", "Computer", "Total size", "Folder name", "Space free", "Type", "Rating", "Kind", "Perceived type", "Item type", "Name")) {
                                $FileData[$Property] = $Detail -replace ';', ''
                            }
                        }
                    }

                    # Process GPS data separately using System.Drawing
                    $Image = [System.Drawing.Image]::FromFile($Item.Path)
                    Try {
                        $GPSLatitude = $Image.GetPropertyItem(2).Value
                        $GPSLatitudeRef = $Image.GetPropertyItem(1).Value
                        $GPSLongitude = $Image.GetPropertyItem(4).Value
                        $GPSLongitudeRef = $Image.GetPropertyItem(3).Value
                        $GPSAltitude = $Image.GetPropertyItem(6).Value

                        $FileData["Latitude"] = Convert-GPSCoordinate -Coordinate $GPSLatitude -Ref ([System.Text.Encoding]::ASCII.GetString($GPSLatitudeRef))
                        $FileData["Longitude"] = Convert-GPSCoordinate -Coordinate $GPSLongitude -Ref ([System.Text.Encoding]::ASCII.GetString($GPSLongitudeRef))
                        $FileData["Altitude"] = [BitConverter]::ToUInt32($GPSAltitude, 0) / [BitConverter]::ToUInt32($GPSAltitude[4..7], 0)
                    } Catch {
                        Write-Output `n"Error: Unable to retrieve GPS data."
                    }

                    $null = $EXIFData.Add([PSCustomObject]$FileData)
                }
            }
        }
    } Else {
        Write-Output `n"Error: Unable to retrieve folder details."
    }
}

# Function to process folders without recursion
Function Process-Folder-NoRecursion {
    Param (
        [Parameter(Mandatory=$True)]
        [string]$FolderPath
    )

    $Shell = New-Object -ComObject Shell.Application
    $Folder = $Shell.Namespace($FolderPath)

    If ($Folder) {
        ForEach ($Item in $Folder.Items()) {
            If (-not $Item.IsFolder -and $Item.Name -match '\.(jpg|jpeg|raw|tiff)$') {
                Write-Output "Processing file: $Item.Name"
                # Ensure the path is valid
                If ([System.IO.File]::Exists($Item.Path)) {
                    Try {
                        $Image = [System.Drawing.Image]::FromFile($Item.Path)
                        # Retrieve and store all the details
                        $FileData = @{}
                        For ($i = 0; $i -lt 266; $i++) {
                            $Detail = $Folder.GetDetailsOf($Item, $i)
                            If ($Detail) {
                                $Property = $Folder.GetDetailsOf($Folder.Items, $i)
                                # Excludes the following fields since they are redundant, but they can easily be added back if needed
                                If ($Property -notin @("Space used", "Owner", "Folder", "Date accessed", "Link status", "Shared", "Computer", "Total size", "Folder name", "Space free", "Type", "Rating", "Kind", "Perceived type", "Item type", "Name")) {
                                    $FileData[$Property] = $Detail -replace ';', ''
                                }
                            }
                        }

                        # Process GPS data separately using System.Drawing
                        Try {
                            $GPSLatitude = $Image.GetPropertyItem(2).Value
                            $GPSLatitudeRef = $Image.GetPropertyItem(1).Value
                            $GPSLongitude = $Image.GetPropertyItem(4).Value
                            $GPSLongitudeRef = $Image.GetPropertyItem(3).Value
                            $GPSAltitude = $Image.GetPropertyItem(6).Value

                            $FileData["Latitude"] = Convert-GPSCoordinate -Coordinate $GPSLatitude -Ref ([System.Text.Encoding]::ASCII.GetString($GPSLatitudeRef))
                            $FileData["Longitude"] = Convert-GPSCoordinate -Coordinate $GPSLongitude -Ref ([System.Text.Encoding]::ASCII.GetString($GPSLongitudeRef))
                            $FileData["Altitude"] = [BitConverter]::ToUInt32($GPSAltitude, 0) / [BitConverter]::ToUInt32($GPSAltitude[4..7], 0)
                        } Catch {
                            Write-Output "Error: Unable to retrieve GPS data."
                        }

                        $Null = $EXIFData.Add([PSCustomObject]$FileData)
                    } Catch {
                        Write-Output "Error: Unable to retrieve image data."
                    }
                } Else {
                    Write-Output "Error: Invalid file path - $Item.Path"
                }
            }
        }
    } Else {
        Write-Output "Error: Unable to retrieve folder details."
    }
}

# Checks if a file or folder was selected
If (-not [string]::IsNullOrEmpty($SelectedPath)) {
    # Initialize an array to store the EXIF data
    $EXIFData = New-Object System.Collections.ArrayList

    # Processes the selected file or folder based on the user's choice
    If ($FileOrFolderChoice -eq '1') {
        $null = Process-File -FilePath $SelectedPath
    } ElseIf ($RecursionChoice -eq '1') {
        $null = Process-Folder -FolderPath $SelectedPath
    } Else {
        $null = Process-Folder-NoRecursion -FolderPath $SelectedPath
    }

    # Create a SaveFileDialog box for saving the Excel spreadsheet
    $SaveSpreadsheet = New-Object System.Windows.Forms.SaveFileDialog
    $SaveSpreadsheet.Filter = "Excel Files|*.xlsx"
    $SaveSpreadsheet.FileName = "EXIF Data.xlsx"
    $SaveSpreadsheet.Title = "Save the results spreadsheet"

    # Shows the file save dialog box
    Write-Host `n"You will now be prompted to select a save location for the results spreadsheet"
    Start-Sleep -Seconds 3
    $Null = $SaveSpreadsheet.ShowDialog()
    $SaveSpreadsheetPath = $SaveSpreadsheet.FileName

    # Checks if a file path was selected
    If (-not [string]::IsNullOrEmpty($SaveSpreadsheetPath)) {
        # To make the spreadsheet a bit more legible, this defines and sets the column order in the spreadsheet and exports it to an Excel spreadsheet
        If ($EXIFData.Count -gt 0) {
            $Order = @("Filename", "Path", "File location", "File extension", "Size", "Date created", "Date modified", "Date taken", "Dimensions", "Width", "Height", "Camera maker", "Camera model", "Latitude", "Longitude", "Altitude")
            $OrderData = $EXIFData | Select-Object -Property @($Order + ($EXIFData[0].PSObject.Properties.Name | Where-Object { $_ -notin $Order }))

            $OrderData | Export-Excel -Path $SaveSpreadsheetPath -WorksheetName "EXIF Data" -TableStyle Medium9 -TableName "EXIF_Data" -AutoSize
        } Else {
            Write-Output `n"No EXIF data found to save."
        }
    } Else {
        Write-Output `n"No file selected for saving."
    }
} Else {
    Write-Output `n"No file or folder selected."
}

# Conclusion message
Write-Output `n"EXIF data processing has completed"
Write-Output "The EXIF data spreadsheet has been saved to $SaveSpreadsheetPath"
Start-Sleep -Seconds 3
Write-Output `n"This window will close shortly"
Start-Sleep -Seconds 5