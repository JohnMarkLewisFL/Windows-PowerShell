# This script converts text to speech using the built-in Windows text-to-speech tools

# Needed for the file save dialog box to function
Add-Type -AssemblyName System.Windows.Forms

# Check if the user wants to choose a .txt file or type a message
$TXTYN = Read-Host "Do you want to choose a .txt file (Y/N)?"

If ($TXTYN -eq "Y" -or $TXTYN -eq "y") {
    # Show a file dialog box to choose the input .txt file
    $OpenTXTDialogBox = New-Object System.Windows.Forms.OpenFileDialog
    $OpenTXTDialogBox.Filter = "Text files (*.txt)|*.txt"
    $OpenTXTDialogBox.Title = "Choose a .txt file"
    $OpenTXTDialogBox.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    If ($OpenTXTDialogBox.ShowDialog() -eq "OK") {
        # Read the content from the selected .txt file
        $TXTFileOrUserMessage = Get-Content -Path $OpenTXTDialogBox.FileName
    } Else {
        Write-Host "User canceled the operation."
        Return
    }
} Else {
    # Prompt the user to type a message
    $TXTFileOrUserMessage = Read-Host "Please type your message"
}

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
$AudioStream = New-Object System.IO.MemoryStream
$Synthesizer.SetOutputToWaveStream($AudioStream)
$Synthesizer.Speak($TXTFileOrUserMessage)

# Specify the output path for the .wav file
$SaveWAVFileBox = New-Object System.Windows.Forms.SaveFileDialog
$SaveWAVFileBox.Filter = "WAV files (*.wav)|*.wav"
$SaveWAVFileBox.Title = "Save Audio File"
$SaveWAVFileBox.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$SaveWAVFileBox.FileName = ".wav"

If ($SaveWAVFileBox.ShowDialog() -eq "OK") {
    $OutputPath = $SaveWAVFileBox.FileName

    # Save the audio stream as a .wav file
    $Synthesizer.SetOutputToWaveFile($OutputPath)
    $Synthesizer.Speak($TXTFileOrUserMessage)  # Re-speak to save to file

    Write-Host "Audio saved to $OutputPath"
} Else {
    Write-Host "User canceled the operation."
}