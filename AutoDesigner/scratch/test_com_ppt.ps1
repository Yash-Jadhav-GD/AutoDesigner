$ppt = New-Object -ComObject PowerPoint.Application
Write-Host "Presentations count: $($ppt.Presentations.Count)"
for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
    $p = $ppt.Presentations.Item($i)
    Write-Host "Presentation ${i}: $($p.FullName)"
}
