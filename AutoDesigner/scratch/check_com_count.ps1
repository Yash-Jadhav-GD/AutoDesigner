$pptxPath = 'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx'
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptxPath, 1, 0, 0)
Write-Host "PowerPoint sees Slide Count: " $pres.Slides.Count
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    Write-Host "Slide $i shape count: " $pres.Slides.Item($i).Shapes.Count
}
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
