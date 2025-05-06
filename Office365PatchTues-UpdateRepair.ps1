# =========================
# Configurable Variables
# =========================
$TaskName       = "Office365_WeeklyUpdateRepair"
$TaskTime       = "06:00"
$TaskDay        = "TUE"
$FilePath = "C:\Scripts"
$ScriptPath     = "$FilePath\Office365UpdateRepair.ps1"
$LogPath        = "$FilePath\OfficeUpdate.log"


# =========================
# Ensure Script Folder Exists
# =========================
if (-not (Test-Path -Path "$FilePath")) {
    New-Item -ItemType Directory -Path "$FilePath" -Force | Out-Null
}

# =========================
# Create the Office Maintenance Script
# =========================
$ScriptContent = @"
`$Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Add-Content -Path '$LogPath' -Value "`$Time - Starting Office maintenance (channel change, update, and repair)..."

# Set Office Channel to Current
`$C2RClient = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
if (Test-Path `$C2RClient) {
    Start-Process -FilePath `$C2RClient -ArgumentList "/changesetting Channel=Current" -Wait
    Add-Content -Path '$LogPath' -Value "`$Time - Channel set to Current."
    
    # Force Silent Office Update
    Start-Process -FilePath `$C2RClient -ArgumentList "/update user displaylevel=false forceappshutdown=true" -Wait
    Add-Content -Path '$LogPath' -Value "`$Time - Office update completed."
} else {
    Add-Content -Path '$LogPath' -Value "`$Time - OfficeC2RClient.exe not found."
}

# Quick Repair
`$RepairPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path `$RepairPath) {
    Start-Process -FilePath `$RepairPath -ArgumentList "scenario=Repair platform=x64 culture=en-us forceappshutdown=True RepairType=QuickRepair DisplayLevel=False" -Wait
    Add-Content -Path '$LogPath' -Value "`$Time - Quick repair completed."
} else {
    Add-Content -Path '$LogPath' -Value "`$Time - OfficeClickToRun.exe not found."
}
"@
$ScriptContent | Set-Content -Path $ScriptPath -Force -Encoding UTF8

# =========================
# Create Scheduled Task (SYSTEM)
# =========================
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $TaskDay -At $TaskTime
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Description "Weekly Office 365: Set Channel, Update, and Quick Repair" -Force

Write-Host "✅ Scheduled task '$TaskName' created to run every $TaskDay at $TaskTime as SYSTEM."
