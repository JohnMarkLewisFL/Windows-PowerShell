# This script generates a random number with user-specified options

Do {
    # Asks the user if they want to use a user-defined range or the default range of 1 to 2147483647
    $YesNoRange = Read-Host "Do you want to use a user-defined range? (Y/N or Yes/No)"
    $YesNoRange = $YesNoRange.ToLower()

    If ($YesNoRange -eq "y" -or $YesNoRange -eq "yes") {
        # Asks the user for the minimum and maximum values
        Do {
            $Min = Read-Host "Enter the minimum number (>=1)"
            $Max = Read-Host "Enter the maximum number (<=2147483647)"
            If ($Min -match "^[0-9]+$" -and $Max -match "^[0-9]+$") {
                $Min = [int]$Min
                $Max = [int]$Max
                If ($Min -ge 1 -and $Max -le 2147483647 -and $Min -lt $Max) {
                    Break
                } Else {
                    Write-Host "Invalid input. The minimum number should be >=1 and the maximum number should be <=2147483647 while being >minimum."
                }
            } Else {
                Write-Host "Invalid input. Please enter numeric values between 1 and 2147483647."
            }
        } While ($true)
    } Else {
        # Sets the minimum and maximum values for the random number generator
        $Min = 1
        $Max = 2147483647
    }

    # Generates a random number with the user-specified min and max values
    $RandomNumber = Get-Random -Min $Min -Max $Max

    # Outputs the random number
    Write-Host $RandomNumber

    # Asks the user if they want to generate another random number
    $RunAgain = Read-Host "Do you want to generate another random number? (Y/N or Yes/No)"

    # Convert the user input to lowercase for case insensitive comparison
    $RunAgain = $RunAgain.ToLower()

} While ($RunAgain -eq "y" -or $RunAgain -eq "yes")

# If the user input is not 'y' or 'yes', the script will stop generating random numbers and close