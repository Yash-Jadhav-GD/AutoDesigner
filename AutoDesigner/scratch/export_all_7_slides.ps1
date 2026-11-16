$pptxPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx"
$outDir = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e"

try {
    $pptApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
    $pptApp = New-Object -ComObject PowerPoint.Application
}

$pres = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$slideIdx = 1
foreach ($slide in $pres.Slides) {
    $outPath = Join-Path $outDir "slide$slideIdx.png"
    $slide.Export($outPath, "PNG", 1920, 1080)
    Write-Host "Exported Slide $slideIdx to $outPath"
    $slideIdx++
}

$pres.Close()
Write-Host "Done exporting all 7 slides!"
