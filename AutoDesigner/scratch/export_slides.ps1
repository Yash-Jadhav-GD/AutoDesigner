param(
    [string]$pptxPath = 'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx',
    [string]$outDir = 'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\preview'
)
$pptxPath = (Resolve-Path $pptxPath).Path
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
$outDir = (Resolve-Path $outDir).Path

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptxPath, 1, 0, 0)
$count = $pres.Slides.Count
for ($i = 1; $i -le $count; $i++) {
    $pngPath = Join-Path $outDir ("slide$($i).png")
    $pres.Slides.Item($i).Export($pngPath, "PNG", 1920, 1080)
    Write-Host "Exported slide $($i) to $pngPath"
}
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Done exporting $count slides."
