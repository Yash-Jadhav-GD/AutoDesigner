try {
    $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    Write-Host "Connected to PowerPoint. Presentations count: $($ppt.Presentations.Count)"
    for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
        $p = $ppt.Presentations.Item($i)
        Write-Host "Presentation ${i}: $($p.FullName)"
    }
} catch {
    Write-Host "Could not connect to active PowerPoint: $_"
}
