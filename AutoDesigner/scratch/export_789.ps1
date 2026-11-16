$pptxPath = 'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\Design_tester_new_formatted.pptx'
$outDir = 'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\output\preview'

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptxPath, 1, 0, 0)
for ($i = 7; $i -le 9; $i++) {
    try {
        $pngPath = Join-Path $outDir ("slide$($i).png")
        $pres.Slides.Item($i).Export($pngPath, "PNG", 1920, 1080)
        Write-Host "Exported slide $($i) to $pngPath"
    } catch {
        Write-Host "Error on slide $($i): $($_.Exception.Message)"
    }
}
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
