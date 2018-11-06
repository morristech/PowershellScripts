$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

foreach($Mailbox in $Mailboxes) {

    $calendars = Get-MailboxFolderStatistics $Mailbox.UserPrincipalName -FolderScope Calendar

    $identity = $Mailbox.UserPrincipalName + ":\" + $calendars[0].Name

	write-host $identity
	Set-MailboxFolderPermission -Identity $identity -User "Default" -AccessRights "Reviewer"
}