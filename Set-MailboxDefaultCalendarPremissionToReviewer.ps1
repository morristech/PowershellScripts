$mailbox = Read-Host 'Mailbox Name ?' 
$identity = $mailbox + ":\Calendar"
Set-MailboxFolderPermission -Identity $identity -User "Default" -AccessRights "Reviewer"