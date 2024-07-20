# This script checks user-provided stock symbols once every 60 seconds using the PSYahooFinance module


# Import the PSYahooFinance module
# If this module is not installed, run the following cmddlet (may need to be run as administrator):
# Install-Module PSYahooFinance -Force
Import-Module PSYahooFinance

# Function to retrieve the stock quotes
Function Get-StockQuote {
    Param (
        [String]$Symbol
    )
    $Quote = Get-YFStockQuote -Symbol $Symbol.ToUpper() -Interval 1m -Range 1d
    # Returns only the most recent 1 minute quote/candle
    Return $Quote | Select-Object -Last 1
}

# Function to retrieve user-provided stock symbols
Function Get-StockSymbols {
    $Symbols = @()
    Do {
        $Symbol = Read-Host "Enter a stock symbol"
        # Makes the stock symbol uppercase for legibility/presentation
        $Symbols += $Symbol.ToUpper()
        Do {
            $AddMore = Read-Host "Do you want to add another symbol? (yes/no)"
        # Accepts y for yes and n for no
        } While ($AddMore -notmatch '^(yes|y|no|n)$')
    } While ($AddMore -match '^(yes|y)$')
    Return $Symbols
}

# Retrieve the user-provided stock symbols
$StockSymbols = Get-StockSymbols

# Loop to update the stocks table once every 60 seconds
While ($True) {
    $AllQuotes = @()
    Foreach ($Symbol in $StockSymbols) {
        $Quote = Get-StockQuote -symbol $Symbol
        $AllQuotes += $Quote
    }
    Clear-Host
    Write-Host "This table will refresh automatically every 60 seconds with the most recent 1 minute quote:`n`n"
    # Sorts the table to show symbols in ascending alphabetical order
    $AllQuotes | Sort-Object Symbol | Format-Table
    Write-Host "`n`n Press CTRL+C to exit"
    Start-Sleep -Seconds 60
}