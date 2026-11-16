$pptxPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx"
$outDir = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1 # ppAlertsNone
try {
    $pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $slideCount = $pres.Slides.Count
    Write-Host "Exporting $slideCount slides..."
    for ($i = 1; $i -le $slideCount; $i++) {
        $outPath = Join-Path $outDir "formatted_12_slide$i.png"
        $pres.Slides.Item($i).Export($outPath, "PNG", 1920, 1080)
        Write-Host "Exported slide $i to $outPath"
    }
    $pres.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pres.Close()
    Write-Host "Presentation closed cleanly!"
} finally {
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
Write-Host "Export completed successfully!"
