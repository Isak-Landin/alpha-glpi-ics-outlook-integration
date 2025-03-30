# Define URL with your generated token
$PersonalToken = "ov9ZA0NpTZoRytgbWhl7IBSdX77CLTcsNDZAA1QM" 

# Comparison URL
# $icsUrl = "https://support.compliq.se/front/planning.php?genical=1&uID=306&gID=0&entities_id=0&is_recursive=1&token=0sG8nv7V0owVbBlWwx5tqPUmT3Blu0PEZoqQPZbz"
$icsUrl = "https://support.compliq.se/front/planning.php?genical=1&uID=306&gID=0&entities_id=0&is_recursive=1&token=0sG8nv7V0owVbBlWwx5tqPUmT3Blu0PEZoqQPZbz"

## Personal token not working, just generating a blank page
# $icsUrl = "https://support.compliq.se/front/planning.php?genical=1&uID=306&gID=0&entities_id=0&is_recursive=1&token=ov9ZA0NpTZoRytgbWhl7IBSdX77CLTcsNDZAA1QM"

Write-Host $icsUrl

$customPath = "$env:USERPROFILE\glpiToOutlook"

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
$downloadPath = $customPath + "\calendar_direct.ics"

# Download the ICS file
# TODO This part needs to be implemented when we know how to save request as repsonse in ps1

Invoke-WebRequest -Uri $icsUrl -OutFile $downloadPath

Write-Host "ICS file downloaded successfully: $downloadPath" -ForegroundColor Green
