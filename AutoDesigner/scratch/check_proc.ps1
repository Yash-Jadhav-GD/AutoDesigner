Get-CimInstance Win32_Process -Filter "Name = 'POWERPNT.EXE'" | Select-Object ProcessId, CommandLine, CreationDate | Format-List
