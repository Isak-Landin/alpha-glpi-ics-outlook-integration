# User Configuration Path
$ScriptDirectory = Split-Path $ScriptPath -Parent

# Load config from psd1
$config = Import-PowerShellDataFile "$PSScriptRoot\..\data.psd1"

# Comparison URL
$icsUrl = $config.IcsUrl

Write-Host $icsUrl

$customPath = Join-Path "$env:USERPROFILE" "glpiToOutlook"

if (-Not (Test-Path -Path $customPath)){
    # Test if it is the user path or the glpi folder that does not exist
    if (-Not (Test-Path -Path $env:USERPROFILE)){
        Write-Host "For whatever reason, we cannot find USERPROFILE"
        exit 1
    }
    elseif (-Not (Test-Path -Path $customPath)) {
        try:
            mkdir $customPath
            Write-Host "Created $customPath"
        catch:
            Write-Host"Failed to create $customPath even though parent folder exists"
            exit 1
    }
    else {
        Write-Host "Could not find the path but we do not know why. For safety reasons, we cannot continue"
        exit 1
    }
}

# Define where to save the ICS file
$downloadPath = Join-Path $customPath "calendar_direct.ics"

# Download the ICS file
# TODO This part needs to be implemented when we know how to save request as repsonse in ps1

Invoke-WebRequest -Uri $icsUrl -OutFile $downloadPath

Write-Host "ICS file downloaded successfully: $downloadPath" -ForegroundColor Green
