$pptxPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx"
$outPath = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e\fresh_s12_test.png"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1 # ppAlertsNone
try {
    $pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $pres.Slides.Item(12).Export($outPath, "PNG", 1920, 1080)
    Write-Host "Exported Slide 12 to $outPath"
    $pres.Close()
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
Write-Host "Done!"
