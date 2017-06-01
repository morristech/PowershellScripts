## Credentials
$UserAdmin = ""
$Credentials = Get-Credential -Credential $UserAdmin

## Connection into Office 365 Management
Connect-MsolService -Credential $Credentials
$MsoExchangeURL = "https://ps.outlook.com/PowerShell-LiveID?PSVersion=5.0.10586.122"

## Connection into Exchange Management
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $MsoExchangeURL -Credential $Credentials -Authentication Basic -AllowRedirection

## Import session allowing to override current commands
Import-PSSession $Session -AllowClobber