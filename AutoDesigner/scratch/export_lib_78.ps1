$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open("C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Library_inspect.pptx", [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.Slides.Item(7).Export("C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e\library_slide7.png", "PNG", 1920, 1080)
$pres.Slides.Item(8).Export("C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e\library_slide8.png", "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Exported library slides 7 and 8."
