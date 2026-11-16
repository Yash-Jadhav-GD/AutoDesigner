$pptxPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\scratch_test_bullet.pptx"
$outPath = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e\test_bullet.png"
try {
    $pptApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
    $pptApp = New-Object -ComObject PowerPoint.Application
}
$pres = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.Slides[1].Export($outPath, "PNG", 1920, 1080)
$pres.Close()
Write-Host "Exported test_bullet.png successfully!"
