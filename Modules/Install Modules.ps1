# This script will automatically install the list of PowerShell modules that I usually use. Please note it must run as Administrator to install the modules properly.

# Elevates the script to run as Administrator
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandpath`" `"$args`"" -Verb RunAs; exit }

# List of modules to install
$Modules = @("ImportExcel", "PSWritePDF", "PSWriteWord", "PSWordCloud", "PSYahooFinance", "PoshInternals", "PowerSploit") 

# Function to install each module
Function Install-ModuleList {
    Param (
        [Parameter(Mandatory=$True)]
        [String]$ModuleName,
        [String]$RequiredVersion,
        [Bool]$AllowClobber = $false
    )

    If (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module $ModuleName"
        If ($RequiredVersion) {
            Install-Module -Name $ModuleName -RequiredVersion $RequiredVersion -Force -Confirm:$False -AllowClobber:$AllowClobber
        } Else {
            Install-Module -Name $ModuleName -Force -Confirm:$False -AllowClobber:$AllowClobber
        }
    } Else {
        Write-Host "Module $ModuleName is already installed"
    }
}

# Foreach loop with if statements to install each module and note the options for specific modules when necessary
Foreach ($Module in $Modules) {
    # PowerSploit needs the required version set to 3.0.0.0
    If ($Module -eq "PowerSploit") {
        Install-ModuleList -moduleName $Module -requiredVersion "3.0.0.0"
    # PoshInternals needs the AllowClobber option set
    } Elseif ($Module -eq "PoshInternals") {
        Install-ModuleList -moduleName $Module -allowClobber $True
    } Else {
        Install-ModuleList -moduleName $Module
    }
}
