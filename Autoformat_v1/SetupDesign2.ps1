$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\Library.pptx')
$slide = $pres.Slides.Item(1)

$d1Decor = $null; $d1Header = $null; $d1Text = $null; $divider = $null
foreach ($shape in $slide.Shapes) {
    if ($shape.Name -eq 'Design 1 Decor Box') { $d1Decor = $shape }
    if ($shape.Name -eq 'Design 1 Header Text') { $d1Header = $shape }
    if ($shape.Name -eq 'Design 1 Body Text 1') { $d1Text = $shape }
    if ($shape.Name -eq 'Grey Centre Fade H Divider') { $divider = $shape }
}

$offset = 350
if ($d1Decor) { $s = $d1Decor.Duplicate(); $s.Name = 'Design 2 Decor Box'; $s.Left += $offset }
if ($d1Header) { $s = $d1Header.Duplicate(); $s.Name = 'Design 2 Header Text'; $s.Left += $offset }
if ($d1Text) { 
    $s = $d1Text.Duplicate(); $s.Name = 'Design 2 Body Text 1'; $s.Left += $offset + 40; $s.Width -= 40 
    $icon = $slide.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeOval, $s.Left - 35, $s.Top + 5, 20, 20)
    $icon.Name = 'Design 2 Icon'
    $icon.Fill.ForeColor.RGB = 0x00FF00
}

$pres.Save()
$pres.Close()
