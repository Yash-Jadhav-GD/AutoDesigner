$pptxPath = "c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\After_copy.pptx"
$outDir = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e"

$pptApp = New-Object -ComObject PowerPoint.Application
$pres = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
for ($idx = 1; $idx -le 4; $idx++) {
    $out = Join-Path $outDir ("after_copy_" + $idx + ".png")
    $pres.Slides.Item($idx).Export($out, "PNG", 1920, 1080)
    Write-Host "Exported after slide $idx"
}
$pres.Close()
$pptApp.Quit()
Write-Host "Done!"
