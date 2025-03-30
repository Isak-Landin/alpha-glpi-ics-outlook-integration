# Define the ICS file path
$ICSFile = "C:\Users\Isak\Downloads\20250325095732.ics"

# Get Outlook Application
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Calendar = $Namespace.GetDefaultFolder(9) # 9 = olFolderCalendar

# Check if Outlook is running
if ($null -eq $Outlook) {
    Write-Host "Error: Outlook could not be started." -ForegroundColor Red
    Exit
}

# Ensure the file exists
if (-Not (Test-Path $ICSFile)) {
    Write-Host "Error: ICS file not found at $ICSFile" -ForegroundColor Red
    Exit
}

# Import the ICS file using Start-Process
Start-Process "OUTLOOK.EXE" -ArgumentList "/importprf $ICSFile"

Write-Host "ICS File Imported Successfully!" -ForegroundColor Green
