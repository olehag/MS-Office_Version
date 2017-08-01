Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  ? {$_.displayname -like "Microsoft Office*"} | Select-Object Displayname, DisplayVersion | Format-Table -AutoSize

Write-Host ""
Write-Host "Press any button to continue..." -ForegroundColor Yellow
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

