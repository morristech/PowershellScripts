
$hostname = $(hostname)

Remove-OwaVirtualDirectory -Identity "$hostname\owa (Default Web Site)"
New-OwaVirtualDirectory -WebSiteName "Default Web Site"
Set-owavirtualdirectory -identity "$hostname\owa (Default Web Site)" -AdfsAuthentication $False -BasicAuthentication $true -WindowsAuthentication $false -DigestAuthentication $false -FormsAuthentication $true -LogonFormat PrincipalName

IISRESET /noforce