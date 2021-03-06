#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;

$Mailboxes = Get-Mailbox 
$ExcludedUsers = @("NT AUTHORITY\SELF","NT AUTHORITY\SYSTEM","NT AUTHORITY\NETWORK SERVICE")
$MailboxPermissionReport = @();

foreach ($Mailbox in $Mailboxes){

    $MailboxACL = Get-MailboxPermission -Identity $Mailbox

    foreach ($MailboxAccess in $MailboxACL){
    
        if($ExcludedUsers -notcontains $MailboxAccess.User){
    
           $MailboxPermissionReportData = @{
                "Identity" = $MailboxAccess.Identity;
                "User" = $MailboxAccess.User;
                "AccessRights" = $MailboxAccess.AccessRights;                            
           }

           $MailboxPermissionReport += (New-Object -TypeName PSObject -Property $MailboxPermissionReportData)

        }
    
    }

}

$MailboxPermissionReport