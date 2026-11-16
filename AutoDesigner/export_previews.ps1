$pptPath = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx"
$outDir = "C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\preview"
$brainDir = "C:\Users\Yash.jadhav\.gemini\antigravity\brain\a7721b22-f173-4015-9588-8200b2c01a9e"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$slideCount = $pres.Slides.Count
Write-Host "Presentation opened. Slide count: $slideCount"

for ($i = 1; $i -le $slideCount; $i++) {
    $outPng = Join-Path $outDir "slide$i.png"
    $brainPng = Join-Path $brainDir "slide$i.png"
    $pres.Slides.Item($i).Export($outPng, "PNG", 1920, 1080)
    Copy-Item $outPng -Destination $brainPng -Force
    Write-Host "Exported and copied slide $i to $outPng and $brainPng"
}

# Also clean up obsolete slide8, slide9 from brain and preview if they exist
8..9 | ForEach-Object {
    $p1 = Join-Path $outDir "slide$_.png"
    $p2 = Join-Path $brainDir "slide$_.png"
    if (Test-Path $p1) { Remove-Item $p1 -Force }
    if (Test-Path $p2) { Remove-Item $p2 -Force }
}

$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Done exporting all slides."
