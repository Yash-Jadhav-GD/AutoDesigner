Get-CimInstance Win32_Process -Filter "Name = 'POWERPNT.EXE'" | Where-Object { $_.CommandLine -like "*AUTOMATION*" } | ForEach-Object {
    Write-Host "Stopping orphaned automation process $($_.ProcessId)"
    Stop-Process -Id $_.ProcessId -Force
}
