# This script converts text to speech using the built-in Windows text-to-speech tools

Do {
    # Prompt the user to type a message
    $UserMessage = Read-Host "Please type your message"

    # Create a SpeechSynthesizer instance
    Add-Type -AssemblyName System.Speech
    $Synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer

    # Ask the user to select the voice
    Do {
        $VoiceInput = Read-Host "Please enter the voice. Type 'm' or 'male' for male voice or 'f' or 'female' for female voice."
        If ($VoiceInput -match "^(m|male|M|Male)$") {
            $Synthesizer.SelectVoice("Microsoft David Desktop")
            $ValidInput = $true
        } ElseIf ($VoiceInput -match "^(f|female|F|Female)$") {
            $Synthesizer.SelectVoice("Microsoft Zira Desktop")
            $ValidInput = $true
        } Else {
            Write-Host "Invalid input. Please enter 'm' or 'male' for male voice or 'f' or 'female' for female voice."
            $ValidInput = $false
        }
    } While ($ValidInput -eq $false)

    # Ask the user to select the speed of the speech
    Do {
        $SpeechRateInput = Read-Host "Please enter the speed of the speech. 0 is normal, 10 is the fastest, and -10 is the slowest."
        Try {
            $SpeechRate = [int]$SpeechRateInput
            If ($SpeechRate -ge -10 -and $SpeechRate -le 10) {
                $ValidInput = $true
            } Else {
                Write-Host "Invalid input. Please enter an integer between -10 and 10."
                $ValidInput = $false
            }
        } Catch {
            Write-Host "Invalid input. Please enter an integer between -10 and 10."
            $ValidInput = $false
        }
    } While ($ValidInput -eq $false)

    $Synthesizer.Rate = $SpeechRate

    # Generate the audio stream
    $Synthesizer.Speak($UserMessage)

    # Prompt the user to continue or exit
    Do {
        $UserInput = Read-Host "Do you want to continue? (Y/N)"
        If ($UserInput -match "^(y|yes)$") {
            $Continue = $true
            $ValidInput = $true
        } ElseIf ($UserInput -match "^(n|no)$") {
            $Continue = $false
            $ValidInput = $true
        } Else {
            Write-Host "Invalid input. Please enter Y or N."
            $ValidInput = $false
        }
    } While ($ValidInput -eq $false)
} While ($Continue)