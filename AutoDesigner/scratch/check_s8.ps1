$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx', 1, 0, 0)
$s8 = $pres.Slides.Item(8)
foreach ($sh in $s8.Shapes) {
    Write-Host "$($sh.Name): Top=$($sh.Top), Height=$($sh.Height), Bottom=$($sh.Top + $sh.Height)"
}
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
