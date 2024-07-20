# This script retrieves a city's weather forecast using the wttr.in API

# Introductory prompt
Write-Host "This script retrieves weather data using the wttr.in API. It requires an active internet connection.`n`n"

# Prompts the user to enter their city
$City = Read-Host -Prompt 'Enter your city'
# Capitalizes the first letter of the city
$City = $City.Substring(0,1).ToUpper()+$City.Substring(1).ToLower()

# List of valid US state abbreviations
$ValidStates = @("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY")

# Loops until a valid state abbreviation is entered
Do {
    # Prompt the user to enter their two-letter state abbreviation
    $State = Read-Host -Prompt `n'Enter your two-letter state abbreviation'
    # Convert the state abbreviation to uppercase
    $State = $State.ToUpper()

    # Validate the state abbreviation
    If ($State -match '^[A-Z]{2}$' -and $ValidStates -contains $State) {
        # Define the URL for the current weather
        $URL_Current = "https://wttr.in/$($City),$($State)"

        # Use Invoke-RestMethod to send a GET request to the service
        $Response_Current = Invoke-RestMethod -Uri $URL_Current

        # Outputs the current weather forecast
        Write-Output `n`n$Response_Current

        # Defines the URL for the 3-day forecast
        $URL_Forecast = "https://wttr.in/$($City),$($State)?format=3"

        # Use Invoke-RestMethod to send a GET request to the service
        $Response_Forecast = Invoke-RestMethod -Uri $URL_Forecast

        # Outputs the 3-day forecast
        Write-Output "The 3-day forecast for $City, $State is: $Response_Forecast"
    } Else {
        Write-Output "Invalid state abbreviation. Please enter a valid two-character state abbreviation."
    }
} While ($State -notmatch '^[A-Z]{2}$' -or $ValidStates -notcontains $State)

# Pauses the script so the user can read the forecast data in the PowerShell window
Pause