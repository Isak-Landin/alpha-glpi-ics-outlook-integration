# Alpha working with specific calendar

# Resolve script path and supporting files
$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDirectory = Split-Path $ScriptPath -Parent
$DependenciesDirectory = $ScriptDirectory + "\" + "dependencies"
$DownloadScript = $DependenciesDirectory + "\" + "tokent_auth_direct_download.ps1"

Write-Host "Downloading new ICS file..."
& $DownloadScript

# Define the ICS file path
$ICSFileName = "calendar_direct.ics"
$ICSFilePath = "$env:USERPROFILE\glpiToOutlook"
$ICSFile = $ICSFilePath + "\" + $ICSFileName

# Define expected calendar name
$ExpectedCalendarName = "TestCreation"

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

# Search for the expected calendar in subfolders
$TargetCalendar = $null
foreach ($Folder in $DefaultCalendar.Folders) {
    Write-Host $Folder.Name
    if ($Folder.Name -eq $ExpectedCalendarName) {
        $TargetCalendar = $Folder
        break
    }
}

# If calendar is still not found, list all available sub-calendars
if ($null -eq $TargetCalendar) {
    Write-Host "Available calendars under 'My Calendars':" -ForegroundColor Yellow
    foreach ($Folder in $DefaultCalendar.Folders) {
        Write-Host "- $($Folder.Name)"
    }
    Write-Host "Error: Could not find the '$ExpectedCalendarName' calendar." -ForegroundColor Red
    Exit
}

# 🛡️ Safety check before deleting
if ($TargetCalendar.Name -ne $ExpectedCalendarName) {
    Write-Host "❌ ERROR: Calendar folder is not '$ExpectedCalendarName'. Aborting deletion!" -ForegroundColor Red
    Exit 1
}

# 🧹 Delete all existing events in the selected calendar
$Items = $TargetCalendar.Items
$Items.Sort("[Start]")
$Items.IncludeRecurrences = $true

$ToDelete = @()
foreach ($Item in $Items) {
    $ToDelete += $Item
}

foreach ($Item in $ToDelete) {
    try {
        $Item.Delete()
    } catch {
        Write-Host "⚠ Failed to delete item: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host "✅ Cleared all existing events from '$ExpectedCalendarName' calendar." -ForegroundColor Cyan

# Read ICS file content
$ICSContent = Get-Content -Path $ICSFile -Raw

# Split into individual VEVENT blocks
$Events = [regex]::Matches($ICSContent, "BEGIN:VEVENT(.*?)END:VEVENT", [System.Text.RegularExpressions.RegexOptions]::Singleline)

foreach ($Event in $Events) {
    $EventBlock = $Event.Groups[1].Value

    # Extract details
    $Subject = [regex]::Match($EventBlock, "SUMMARY:(.*)").Groups[1].Value.Trim()
    $StartTime = [regex]::Match($EventBlock, "DTSTART:(\d{8}T\d{6})").Groups[1].Value
    $EndTime = [regex]::Match($EventBlock, "DTEND:(\d{8}T\d{6})").Groups[1].Value
    $Description = [regex]::Match($EventBlock, "DESCRIPTION:(.*)").Groups[1].Value.Trim()

    # Convert time format
    $StartDateTime = [datetime]::ParseExact($StartTime, "yyyyMMddTHHmmss", $null)
    $EndDateTime = [datetime]::ParseExact($EndTime, "yyyyMMddTHHmmss", $null)

    # Create a new appointment
    $Appointment = $TargetCalendar.Items.Add(1) # 1 = olAppointmentItem
    $Appointment.Subject = $Subject
    $Appointment.Start = $StartDateTime
    $Appointment.End = $EndDateTime
    $Appointment.Body = $Description
    $Appointment.ReminderMinutesBeforeStart = 15
    $Appointment.Save()

    Write-Host "✅ Imported: $Subject from $StartDateTime to $EndDateTime" -ForegroundColor Green
}
