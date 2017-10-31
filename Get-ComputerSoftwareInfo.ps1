$SoftwareName = Read-Host -Prompt "Enter Software Name"
$Software =  Get-WmiObject -Class Win32_Product |  Where-Object {$_.Name -Match $SoftwareName}
Write-Host $Software.Name
Write-Host $Software.Version 