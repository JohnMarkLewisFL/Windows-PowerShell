# This script is designed to make it look like you're doing something really important to non-IT people

# Defines the list of characters to select randomly
$CharacterList = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
                '0','1','2','3','4','5','6','7','8','9',
                '!','@','#','$','%','^','&','*','(',')','-','_','=','+','[',']','{','}','|','\',';',':','"',"'",'<','>',',','.','?','/'," ")

# Function to generate a random character
Function Get-RandomCharacter {
    Return $CharacterList | Get-Random
}

# Function to generate a random number of spaces
Function Get-RandomSpaces {
    $SpaceCount = Get-Random -Minimum 1 -Maximum 5
    Return " " * $SpaceCount
}

# Function to format the random characters in green so it looks somewhat like The Matrix
Function Start-KmartMatrix {
    While ($True) {
        $Width = Get-Random -Minimum 5 -Maximum 75
        $Output = ""
        For ($i = 0; $i -lt $Width; $i++) {
            $Output += Get-RandomCharacter
            $Output += Get-RandomSpaces
        }
        Write-Host $Output -ForegroundColor Green
        Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 75)
    }
}

# Runs the KmartMatrix function indefinitely
Start-KmartMatrix