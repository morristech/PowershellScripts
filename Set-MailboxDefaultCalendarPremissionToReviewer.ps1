$mailbox = Read-Host 'Mailbox Name ?' 
if(Get-Mailbox -Identity $mailbox) {
    Set-MailboxFolderPermission -Identity '$mailbox:\Calendar' -User "Default" -AccessRights "Reviewer"
}
else {
    Write-Host 'No mailbox with such name : $mailbox'
}