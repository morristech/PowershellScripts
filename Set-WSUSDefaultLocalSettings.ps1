# Stop the Windows Update service
Stop-Service -Name wuauserv
# Remove the registry key
Remove-Item HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate -Recurse
# Start the Windows Update service
Start-Service -name wuauserv