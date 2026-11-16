$pptxPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\test_clean_12.pptx"
$outPath = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e\clean_12_exported.png"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1
try {
    $pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $pres.Slides.Item(12).Export($outPath, "PNG", 1920, 1080)
    Write-Host "Exported Slide 12 to $outPath"
    $pres.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
Write-Host "Done!"
