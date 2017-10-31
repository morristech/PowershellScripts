# Source : https://technet.microsoft.com/en-us/library/dn568015.aspx

Remove-PSSession $sfboSession
Remove-PSSession $exchangeSession
Remove-PSSession $ccSession
Disconnect-SPOService