Import-Module ActiveDirectory

Get-ADGroup -filter "Groupcategory -eq 'Security' -AND GroupScope -ne 'DomainLocal' -AND Member -like '*'" |
foreach { 
 Write-Host "Exporting $($_.name)" -ForegroundColor Cyan
 $name = $_.name -replace " ","-"
 $file = Join-Path -path "C:\temp" -ChildPath "$name.csv"
 Get-ADGroupMember -Identity $_.distinguishedname -Recursive |  
 Get-ADObject -Properties SamAccountname,Title,Department |
 Select Name,SamAccountName,Title,Department,DistinguishedName,ObjectClass |
 Export-Csv -Path $file -NoTypeInformation
}