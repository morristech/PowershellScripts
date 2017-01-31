# **** Get Current Folder
#Powershell2
#$CurrentFolderPath = split-path -parent $MyInvocation.MyCommand.Definition

#Powershell3
$CurrentFolderPath = $PSScriptRoot


# **** Import Global Settings
Get-Content $CurrentFolderPath"\GlobalSettings.ini" | foreach-object -begin {$GlobalSettings=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $GlobalSettings.Add($k[0], $k[1]) } }

# **** Initialize Script
#. $PSScriptRoot\Bootstrap.ps1


# **** Get user and Export to CSV file in OutputFolder in Global Settings ini file
Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName" | Where-Object {$_.DisplayName -ne $null} | Select Name, Surname, SamAccountName, DistinguishedName | Export-Csv $($GlobalSettings.Get_Item("OutputFolder") + "\ActiveUserList " + $(get-date -f yyyy-MM-dd) + ".csv") -notype