## Credentials
$UserAdmin = ""
$Credentials = Get-Credential -Credential $UserAdmin

## Connection into Office 365 Management
Connect-MsolService -Credential $Credentials

## Connection into Exchange Management
$MsoExchangeURL = "https://outlook.office365.com/powershell-liveid/"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $MsoExchangeURL -Credential $Credentials -Authentication Basic -AllowRedirection

## Import session allowing to override current commands
Import-PSSession $Session -AllowClobber