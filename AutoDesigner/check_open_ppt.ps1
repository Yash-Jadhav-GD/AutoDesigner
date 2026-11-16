$ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
    $pres = $ppt.Presentations.Item($i)
    Write-Host "Open presentation: $($pres.FullName)"
}
