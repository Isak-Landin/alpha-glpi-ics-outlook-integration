# Define login credentials and URLs
$loginUrl = "https://support.compliq.se"
$icsUrl = "https://support.compliq.se/caldav.php/calendars/users/il@compliq.se/calendar.ics"
$sessionFile = "$env:TEMP\session_cookies.txt"
$downloadPath = "C:\Users\Isak\Downloads\calendar.ics"

# Define login form data (Modify if needed)
$loginData = @{
    "login_name" = "il@compliq.se"
    "login_password" = "your_actual_password_here"  # Do not hardcode; use secure storage!
}

# Step 1: Authenticate & Save Cookies
Invoke-WebRequest -Uri $loginUrl -Method Post -SessionVariable session -Body $loginData

# Step 2: Download the ICS file using session cookies
Invoke-WebRequest -Uri $icsUrl -WebSession $session -OutFile $downloadPath

Write-Host "ICS file downloaded successfully: $downloadPath" -ForegroundColor Green
