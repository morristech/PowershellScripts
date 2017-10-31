$mailbox = Read-Host 'Mailbox Name ?' 
Get-Mailbox $mailbox | select -expandproperty GrantSendOnBehalfTo