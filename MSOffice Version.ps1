#Lord Hagen / olehag04@nfk.no

#Get list of installed items from Uninstall registery. Get only those with displayname "Like" 'Microsoft Office*'. Get object DisplayName, DisplayVersion. Format output -Autosize.
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  Where-Object {$_.displayname -like "Microsoft Office*"} | Select-Object DisplayName, DisplayVersion | Format-Table -AutoSize

#Press a button to exit
Write-Host ""
Write-Host "Press any button to exit..." -ForegroundColor Yellow
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
