$mailbox = Read-Host 'Mailbox Name ?' 

$calendars = Get-MailboxFolderStatistics $mailbox -FolderScope Calendar

$identity = $mailbox + ":\" + $calendars[0].Name

Set-MailboxFolderPermission -Identity $identity -User "Default" -AccessRights "Reviewer"