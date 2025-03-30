# Define the ICS file path
$ICSFile = "C:\Users\Isak\Downloads\20250325095732.ics"

# Get Outlook Application
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Calendar = $Namespace.GetDefaultFolder(9) # 9 = olFolderCalendar

$test_path_result = Test-Path $ICSFile
Write-Host "Test Path result: $test_path_result"
exit 0


try{
    # Import ICS file
    $Calendar.Items.Add($ICSFile)
    }
catch{
    $_
    Write-Host "Could not complete"
    exit 1
    }

Write-Host "ICS File Imported Successfully"
