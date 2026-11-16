$pptxPath = "c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Library_copy.pptx"
$outDir = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e"

$pptApp = New-Object -ComObject PowerPoint.Application
$pres = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$idx = 1
foreach ($s in $pres.Slides) {
    $out = Join-Path $outDir ("library_all_" + $idx + ".png")
    $s.Export($out, "PNG", 1920, 1080)
    Write-Host "Exported library slide $idx"
    $idx++
}
$pres.Close()
$pptApp.Quit()
Write-Host "Done!"
