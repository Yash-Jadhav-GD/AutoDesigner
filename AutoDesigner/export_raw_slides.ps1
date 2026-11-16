$ppt = New-Object -ComObject PowerPoint.Application
$p = (Resolve-Path 'AutoDesigner\Design_tester_new_inspect.pptx').Path
$pres = $ppt.Presentations.Open($p, -1, 0, 0)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $outPath = Join-Path (Get-Location) ('AutoDesigner\output\input_slide' + $i + '.png')
    $pres.Slides.Item($i).Export($outPath, 'PNG', 1920, 1080)
    Write-Host "Exported input slide $i to $outPath"
}
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
