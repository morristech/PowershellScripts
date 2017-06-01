# Mailbox Premission

Set-MailboxFolderPermission -User "Default" -AccessRights "Reviewer" -Identity MREastAfrica@wsscc.org
Set-MailboxFolderPermission -User "Default" -AccessRights "Reviewer" -Identity MRSouthAsia@wsscc.org
Set-MailboxFolderPermission -User "Default" -AccessRights "Reviewer" -Identity MRSouthEastAsia@wsscc.org
Set-MailboxFolderPermission -User "Default" -AccessRights "Reviewer" -Identity MRWestAfrica@wsscc.org

# Calendar Persmission

Set-MailboxFolderPermission -Identity "MREastAfrica@wsscc.org:\Calendar" -User "Default" -AccessRights "Reviewer"
Set-MailboxFolderPermission -Identity "MRSouthAsia@wsscc.org:\Calendar" -User "Default" -AccessRights "Reviewer"
Set-MailboxFolderPermission -Identity "MRSouthEastAsia@wsscc.org:\Calendar" -User "Default" -AccessRights "Reviewer"
Set-MailboxFolderPermission -Identity "MRWestAfrica@wsscc.org:\Calendar" -User "Default" -AccessRights "Reviewer"