$retry = 0
$ppt = $null
while ($retry -lt 5 -and $ppt -eq $null) {
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
    } catch {
        Write-Host "PowerPoint COM busy, retrying in 2s..."
        Start-Sleep -Seconds 2
        $retry++
    }
}

if ($ppt -eq $null) {
    Write-Host "Could not acquire PowerPoint COM object."
    exit 1
}

$pptxPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\test_new_designs.pptx"
$outDir = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e"

$pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$slideIdx = 1
foreach ($slide in $pres.Slides) {
    $outPath = Join-Path $outDir "test_design_slide$slideIdx.png"
    $slide.Export($outPath, "PNG", 1920, 1080)
    Write-Host "Exported Slide $slideIdx to $outPath"
    $slideIdx++
}

$pres.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Done exporting test slides!"
