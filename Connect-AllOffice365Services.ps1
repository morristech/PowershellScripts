# Source : https://technet.microsoft.com/en-us/library/dn568015.aspx

$domainHost = Read-Host 'Domain ?'
$credential = Get-Credential

# Microsoft Online Services
Import-Module MsOnline
Connect-MsolService -Credential $credential

# Sharepoint
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://$domainHost-admin.sharepoint.com -credential $credential

# Skype Online
Import-Module SkypeOnlineConnector
$sfboSession = New-CsOnlineSession -Credential $credential

# Exchange Online
Import-PSSession $sfboSession
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection

Import-PSSession $exchangeSession -DisableNameChecking
$ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection

Import-PSSession $ccSession -Prefix cc