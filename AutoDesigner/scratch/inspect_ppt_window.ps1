Add-Type -AssemblyName UIAutomationClient
$proc = Get-Process POWERPNT -ErrorAction SilentlyContinue
if ($proc) {
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $proc.Id)
    $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)
    foreach ($win in $windows) {
        Write-Host "Window: $($win.Current.Name) | Class: $($win.Current.ClassName)"
        $subChildren = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($c in $subChildren) {
            if ($c.Current.Name) {
                Write-Host "  Child: $($c.Current.Name)"
            }
        }
    }
}
