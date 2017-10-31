$mailbox = Read-Host 'Mailbox Name ?' 
Set-MailboxFolderPermission -Identity '$mailbox:\Calendar' -User "Default" -AccessRights "Reviewer"