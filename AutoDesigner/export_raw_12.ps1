$ErrorActionPreference = 'Stop'
Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$pptPath = 'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design tester new.pptx'
$brainDir = 'C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e'

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $pngPath = Join-Path $brainDir "raw_new_slide$i.png"
    $pres.Slides.Item($i).Export($pngPath, 'PNG', 1920, 1080)
    Write-Host "Exported slide $i to $pngPath"
}

$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host 'Export complete.'
