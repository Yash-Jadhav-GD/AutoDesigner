param (
    [string]$PptxPath,
    [string]$PngPath
)

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($PptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.Slides.Item(1).Export($PngPath, "PNG", 1920, 1080)
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Exported to $PngPath"
