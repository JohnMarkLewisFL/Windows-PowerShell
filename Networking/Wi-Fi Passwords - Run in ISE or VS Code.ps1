# This script will gather all saved Wi-Fi SSIDs along with their respective PSKs


# Introduction message
Write-Host "This script will list all the saved Wi-Fi SSIDs and PSKs for all user profiles.`n`n"
Start-Sleep -Seconds 3

# Gather the SSID from the netsh command and split the unused text and whitespaces
$WirelessProfiles = netsh wlan show profiles | Select-String -Pattern "All User Profile" | %{ ($_ -split ":")[-1].Trim() };

# For each profile found with the netsh command, show the password in plaintext and display the SSID in columns
$WirelessProfiles | foreach {
	$ProfileData = netsh wlan show profiles name=$_ key="clear";
	$SSID = $ProfileData | Select-String -Pattern "SSID Name" | %{ ($_ -split ":")[-1].Trim().Trim('"') };
	$PSK = $ProfileData | Select-String -Pattern "Key Content" | %{ ($_ -split ":")[-1].Trim() };
    $List =	[PSCustomObject]@{
		SSID = $SSID;
		Password = $PSK
	    } 
    Write-Output $List
    Write-Host ""
}
Pause