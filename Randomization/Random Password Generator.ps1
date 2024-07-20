# Import System.Web assembly
Add-Type -AssemblyName System.Web

# Import assembly for clipboard functionality
Add-Type -AssemblyName System.Windows.Forms

# Define the characters to exclude
$ExcludeChars = '`', '~', '-', '_', '=', '+', '[', '{', ']', '}', '\', '|', ';', ':', "'", '"', ',', '<', '.', '/'

Do {
    # Defines the password length
    $Length = Read-Host "Pick a number (preferably between 8 and 20) for the password length" 

    # Defines the number of special characters
    $NumberOfSpecialCharacters = Read-Host "Pick a number (preferably between 1 and 6) for the number of special characters" 

    # Generate random password
    Do {
        $RandomPassword = [System.Web.Security.Membership]::GeneratePassword($Length,$NumberOfSpecialCharacters)
    } While ($RandomPassword.IndexOfAny($ExcludeChars.ToCharArray()) -ne -1)

    # Outputs the random password
    Write-Host "The following password has been generated and is copied to your clipboard:" $RandomPassword

    # Copy the password to the clipboard
    [System.Windows.Forms.Clipboard]::SetText($RandomPassword)

    # Ask the user if they want to generate another password
    $RunAgain = Read-Host "Do you want to generate another password? (y/n)"
} While ($RunAgain -in @("y", "Y", "yes", "Yes"))