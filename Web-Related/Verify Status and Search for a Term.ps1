# This script uses the Invoke-WebRequest cmdlet to verify if a website is online and subsequently search for a word on that page


# Starts a Do While loop to re-run the script if necessary
Do {
    # Prompts the user to enter a URL
    $URL = Read-Host `n"Please enter the URL of the website you are trying to reach"

    # Checks if the URL is online
    Try {
        $URLResponse = Invoke-WebRequest -Uri $URL -UseBasicParsing
        Write-Host `n"$URL is online."
    } Catch {
        Write-Host `n"$URL is not reachable."
        Exit
    }

    # Prompts the user to enter a search term
    $SearchTerm = Read-Host `n"Please enter a search term"

    # Searches for the search term on the URL if the URL is online
    If ($URLResponse.Content -Match $SearchTerm) {
        Write-Host `n"The search term '$SearchTerm' was found on $URL"
    } Else {
        Write-Host `n"The search term '$SearchTerm' was not found on $URL"
    }

    # Prompts the user to run the script again
    $RunAgain = Read-Host `n"Do you want to run the script again? (y/n)"
} While ($RunAgain -Match '^(?i)y(?:es)?$')
