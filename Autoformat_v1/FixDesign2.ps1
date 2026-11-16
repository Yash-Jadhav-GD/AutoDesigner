$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\Library.pptx')
$slide = $pres.Slides.Item(1)

$d2Decor = $null; $d2Header = $null; $d2Text = $null; $d2Icon = $null
foreach ($shape in $slide.Shapes) {
    if ($shape.Name -eq 'Design 2 Decor Box') { $d2Decor = $shape }
    if ($shape.Name -eq 'Design 2 Header Text') { $d2Header = $shape }
    if ($shape.Name -eq 'Design 2 Body Text 1') { $d2Text = $shape }
    if ($shape.Name -eq 'Design 2 Icon') { $d2Icon = $shape }
}

if ($d2Icon) { $d2Icon.Delete() }

if ($d2Decor -and $d2Header) {
    # Match image perfectly: Header is flush with the Decor Box top/left/right
    $d2Header.Left = $d2Decor.Left
    $d2Header.Top = $d2Decor.Top
    $d2Header.Width = $d2Decor.Width
    
    # Ensure it looks like a clean design
    $d2Decor.AutoShapeType = [Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle
}

if ($d2Text -and $d2Decor) {
    # Adjust body text left alignment to match the padding inside the decor box
    $d2Text.Left = $d2Decor.Left + 20
}

$pres.Save()
$pres.Close()
