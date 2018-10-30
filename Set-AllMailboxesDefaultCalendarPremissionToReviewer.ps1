$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

foreach($Mailbox in $Mailboxes) {
	$identity = $Mailbox.Name + ":\Calendar"
	write-host $identity
	Set-MailboxFolderPermission -Identity $identity -User "Default" -AccessRights "Reviewer"
}