#Working alpha

# Define the ICS file path
$ICSFile = "C:\Users\Isak\Downloads\20250325095732.ics"

# Check if the ICS file exists
if (-Not (Test-Path $ICSFile)) {
    Write-Host "Error: ICS file not found at $ICSFile" -ForegroundColor Red
    Exit
}

# Get Outlook Application
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Get the Default Calendar Folder
$DefaultCalendar = $Namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

# Search for "AutoImport" in subfolders
$AutoImportCalendar = $null
foreach ($Folder in $DefaultCalendar.Folders) {
    if ($Folder.Name -eq "AutoImport") {
        $AutoImportCalendar = $Folder
        break
    }
}

# If calendar is still not found, list all available sub-calendars
if ($null -eq $AutoImportCalendar) {
    Write-Host "Available calendars under 'My Calendars':" -ForegroundColor Yellow
    foreach ($Folder in $DefaultCalendar.Folders) {
        Write-Host "- $($Folder.Name)"
    }
    Write-Host "Error: Could not find the 'AutoImport' calendar." -ForegroundColor Red
    Exit
}

# Read ICS file content
$ICSContent = Get-Content -Path $ICSFile -Raw

# Extract event details from ICS file (Basic parsing)
$Subject = [regex]::Match($ICSContent, "SUMMARY:(.*)").Groups[1].Value
$StartTime = [regex]::Match($ICSContent, "DTSTART:(\d{8}T\d{6})").Groups[1].Value
$EndTime = [regex]::Match($ICSContent, "DTEND:(\d{8}T\d{6})").Groups[1].Value

# Convert time format (ICS uses YYYYMMDDTHHMMSSZ format)
$StartDateTime = [datetime]::ParseExact($StartTime, "yyyyMMddTHHmmss", $null)
$EndDateTime = [datetime]::ParseExact($EndTime, "yyyyMMddTHHmmss", $null)

# Create a new calendar appointment in "AutoImport"
$Appointment = $Outlook.CreateItem(1) # 1 = olAppointmentItem
$Appointment.Subject = $Subject
$Appointment.Start = $StartDateTime
$Appointment.End = $EndDateTime
$Appointment.Body = "Imported from ICS"
$Appointment.ReminderMinutesBeforeStart = 15
$Appointment.Move($AutoImportCalendar)  # Move event to "AutoImport" calendar
$Appointment.Save()

Write-Host "ICS Event Imported to 'AutoImport' Calendar: $Subject from $StartDateTime to $EndDateTime" -ForegroundColor Green
