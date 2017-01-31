# **** Get Current Folder
#Powershell2
#$CurrentFolderPath = split-path -parent $MyInvocation.MyCommand.Definition

#Powershell3
$CurrentFolderPath = $PSScriptRoot


# **** Import Global Settings
Get-Content $CurrentFolderPath"\GlobalSettings.ini" | foreach-object -begin {$GlobalSettings=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $GlobalSettings.Add($k[0], $k[1]) } }