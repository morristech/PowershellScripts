#requires -version 5
<#
.SYNOPSIS
    This script will create an Excel report of the Office 365 environment.

    It is neccesarry to use the Microsoft Azure Active Directory Module for Windows PowerShell Preview version 8808.1, released on: 10/30/2015

    Created by Reinout Dorreboom

    The script is based on the template wich is created by Reinout Dorreboom.

.DESCRIPTION
    The script will create several CSV files which will later be imported in an Excel Sheet. Each of these CSV files are later imported in seperate sheets in the workbook. 
    
    The Workbook will consist of the next sheets:
    - AllMSOLUsers - with detailed information of all MSOL users
    - Devices - detailed information about the devices which are registered to Azure AD
    - AvailableLicenses - With available and used licenses of the organization
    - Domains - with all connected domains to Office 365
    - MailContacts - With all Mailcontacts which are entered in the MSOL environment in Exchaneg Online
    - Resource Mailboxes - With all equipment mailboxes (if any)
    - Shared Mailboxes - With all Shared mailboxes including who has access to it and who may send as that particular mailbox
    - SharePointSites - With all SharePoint Sites including used templates, storage quota etc.
    - SkypeForBusiness - Skype for business information, which domains are allowed and wich not
    - UserAndGroups - All groups and memberships
    - UserLicenses - What licenses does each user have
    - UserMailboxes - Detailed information about each mailbox in the MSOL Exchange Online environment.

    Also the workbook consists of pivot tables and will freeze panes for beter readability

    Since Microsoft is using other names for licenses then the names which are commonly used (i.e. EnterPrisePack instead of E3 licenses) a translation table is used
    However it only containce the licenses i have used. If the script errors, please complete the table at line 556.


.PARAMETER
    The script can be run without any parameter. After starting the script it will ask two questions to the user:
    "Where must the report be saved?" and "Do you want to open the report after it has been configured?"

    On line 73 you can provide an accountname of a Office 365 Global Admin. If left empty, the script will ask for it
    On line 74 you can provide the password of the Office 365 Gobal Admin. If left empty, the script will ask for it
    On line 75 you ca provide the URL to the SharePoint Admin site. If left empty, the script will ask for it.



.INPUTS
    None

.OUTPUTS
    A log file will be saved in the temp environment: $env:temp

.NOTES
    Template Version:  1.0   
    Version:           2.4
    Author:            Reinout Dorreboom
    Creation Date:     18 januari 2016
    Purpose/Change:   Initial script development


.PreReqs
    Requirements:
    1.Install the 64-bit version of the Microsoft Online Services Sign-in Assistant  
    2.Install the 64-bit version of the Windows Azure Active Directory Module for Windows PowerShell
       *** It is neccesarry to use the Microsoft Azure Active Directory Module for Windows PowerShell Preview version 8808.1, released on: 10/30/2015 (or higher) ***
    3.Install the 64-bit version of the SharePoint Online Management Shell
    4.Install the Skype for Business Online, Windows PowerShell Module
    5.The script needs to run with Administrative privileges (Run as Admin)   

  
.EXAMPLE
  Create Office Report.ps1
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#clear the Screen
Clear-Host 

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Find open PowerShell sessions and close them
Get-PSSession|Remove-PSSession

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$ScriptVersion = "4.0"

#Provide the credentials, if left empty the script will ask for it
$User 					= ''
$Password 				= ''
$site 					= ''

#Log File Info
$scriptName = $MyInvocation.MyCommand.Name
$startTime = Get-Date -Format 'yyy-MM-dd_HH-mm-ss'
$LogfileName = "$scriptName $startTime.log"
$logfilePath = "$env:temp\"
$Logfile = $logfilePath + $LogfileName



#-----------------------------------------------------------[Functions]------------------------------------------------------------
#region Template Functions
<#

Function <FunctionName>{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Tell what is going on" -type Informational}
  
    Process{
        Try{
        
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfull ended action.." -type Informational }
}


#>

function Test-Administrator{
    $Currentuser = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $Currentuser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}

Function LogWrite{
    Param(
        $logfile = "$logfile",
        [validateset("Informational","Warning","Error")]$type = "Informational",
        [string]$Logstring
    )
  
  Begin{ }
  
  Process{
    Try{
        if($type -eq "Informational"){$foreGroundColor = "Green"}
        if($type -eq "Warning"){$foreGroundColor = "Cyan"}
        if($type -eq "Error"){$foreGroundColor = "Red"}
        Add-content $Logfile -value "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') - $type - $logstring"
  	    Write-Host $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') - $logstring -ForegroundColor $foreGroundColor 
    
    }
    
    Catch{
      Write-Host $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') - $error -ForegroundColor Cyan
      Break
    }
  }
  
  End{ }
}

Function CreateCreds{
    Param(
        $site = "$site",
        $user = "$User",
        $password = "$Password"
    )
  
    Begin{logwrite -Logstring "Creating credentials for user $user" -type Informational }
  
    Process{
        Try{
            if ($site -eq $null -or $site -eq ""){
                logwrite -Logstring "The Site URL is not provided in the script, so it is asked" -type Informational
                $site = Read-Host -Prompt "HTTPS URL for your SP Online admin site (ie. https://[tenantname]-admin.sharepoint.com)"
            }
            if($User -eq $null -or $user -eq ""){
                logwrite -Logstring "The Username is not provided in the script, so it is asked" -type Informational
                $User = Read-Host -Prompt "Username to log in to Office 365"
            } 
            if($password -eq $null -or $password -eq ""){
                logwrite -Logstring "The Password is not provided in the script, so it is asked" -type Informational
                $creds = Get-Credential -UserName $User -Message "Password for $User"
            }
            else {
                $Password = ConvertTo-SecureString $Password -AsPlainText -Force
                $creds = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $User, $PassWord
            }
            return $creds,$site        
        }
    
        Catch{
            logwrite -Logstring $error -type Error
            Break
        }
    }
  
    End{ logwrite -Logstring "The credentials are generated" -type Informational }
}

Function Login2MSOL{
    Param(
        [Parameter(Mandatory=$True)]$creds
    )
  
    Begin{logwrite -Logstring "Signing in to MS Online services" -type Informational}
  
    Process{
        Try{Import-Module MSOnline -DisableNameChecking} Catch{logwrite -Logstring $error -type Error;Break}
        Try{Connect-MsolService –Credential $Creds} Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfully signed in to MS Online services" -type Informational }
}

Function Login2SPO{
    Param(
        [Parameter(Mandatory=$True)]$creds,
        [Parameter(Mandatory=$True)]$site
    )
  
    Begin{logwrite -Logstring "Signing in to SharePoint Online services" -type Informational} 
  
    Process{
        Try{Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking } Catch{logwrite -Logstring $error -type Error;Break}
        Try{Connect-SPOService -Url $site -Credential $Creds} Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{ logwrite -Logstring "Successfully signed in to SharePoint Online Services" -type Informational }
}

Function Login2SfB{
    Param(
        [Parameter(Mandatory=$True)]$creds
    )
  
    Begin{logwrite -Logstring "Signing in to Skype For Business" -type Informational} 
  
    Process{
        Try{$SkypeSession = New-CsOnlineSession -Credential $Creds} Catch{logwrite -Logstring $error -type Error;Break}
        Try{Import-PSSession $SkypeSession -AllowClobber -WarningAction SilentlyContinue -DisableNameChecking |Out-Null} Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{ logwrite -Logstring "Successfully signed in to Skype For Business" -type Informational}
}

Function Login2ExO{
    Param(
        [Parameter(Mandatory=$True)]$creds
    )
  
    Begin{logwrite -Logstring "Signing in to Exchange Online" -type Informational} 
  
    Process{
        Try{$ExOnSession = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue} Catch{logwrite -Logstring $error -type Error;Break}
        Try{Import-PSSession $ExOnSession -WarningAction SilentlyContinue -DisableNameChecking |Out-Null} Catch{logwrite -Logstring $error -type Error ; Break}

    }
  
    End{logwrite -Logstring "Successfully signed in to Exchange Online" -type Informational }
}

Function GetTenantID{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Retrieving the Tenantname" -type Informational}
  
    Process{
        Try{
            #Get the tenant ID
            $sku = Get-MsolAccountSku |select AccountSkuId
            $pos = $sku[0].accountSkuId.indexOf(":")
            $tenant = $sku[0].accountSkuId.ToString()
            return $tenant.Substring(0,$pos)
        } 
        
        Catch{logwrite -Logstring $error -type Error ; Break}
    }
  
    End{logwrite -Logstring "Successfull retrieved the TenantName: $($tenant.Substring(0,$pos))" -type Informational }
}

#endregion Template Functions

Function OpenFileBrowserDialog{
    Param(
        [string]$Description
    )
  
    Begin{logwrite -Logstring "User is asked to provide the location of the Reportfile" -type Informational}
  
    Process{
        Try{   
            Add-Type -AssemblyName System.Windows.Forms
            $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $FolderBrowser.Description = $Description
            [void]$FolderBrowser.ShowDialog()
            return $FolderBrowser.SelectedPath
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "The user wants to save the report file in: $($FolderBrowser.SelectedPath)" -type Informational }
}

Function DisplayMessageBox{
    Param(
        [string]$Message,
        [string]$WindowTitle
    )
  
    Begin{logwrite -Logstring "The user is asked if the Reportfile needs to be opened." -type Informational }
  
    Process{
        Try{
            [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")|Out-Null
            $msgbox = [System.Windows.Forms.MessageBox]::Show($Message,$WindowTitle, 4)
            return $msgbox
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "The user answered $msgbox on the question: $Message" -type Informational }
}

Function WhatToDo{
    Param(
        
    )
  
    Begin{}
  
    Process{
        Try{
            $arrfunctions = @()
            Do{
                Write-host "What kind of report do you want"
                Write-host "A - Create All reports            - Create all reports"
                Write-host "B - Domains                       - Create a report of the registered domains"
                Write-Host "C - Available Licenses            - Create a report of the acquired licenses and availablity of them"
                Write-Host "D - Office 365 Roles              - Create a report of all Office 365 roles and who is assigned to them"
                Write-Host "E - All MSOL Users                - Create a report of all users that are in Azure AD and detailed information about them"
                Write-Host "F - Groups and groupmembers       - Create a report of all groups show all security groups which are in Office 365 (created or synced) with their members"
                Write-Host "G - UserLicenses                  - Create a report of all licenses assigned to each user"
                Write-Host "H - Devices                       - Create a report of all registered devices that are registered in Azure AD"
                Write-Host "I - Mailboxes                     - Create a report of all mailboxes with detailed information about them"
                Write-Host "J - Mailcontacts                  - Create a report of all Mailcontacts with detailed information about them" 
                Write-Host "K - SharePoint sites              - Create a report of all site collections in SharePoint Online"
                Write-Host "L - Office 365 Groups             - Create a report of all Office 365 groups and details"
                Write-Host "M - Skype for Business            - Create a report of all general Skype for business settings "
                Write-Host "N - Skype for Business Usage      - Create a report of the usage of Skype for Business of the last three months on a daily base"
                Write-Host "O - Skype for Business Devices    - Create a report of the used devices for Skype for business"
                Write-Host "P - Skype for Business Activities - Create a report of the usage of Skype for business on a per user base of the last three months"
                Write-Host "X - Exit"

                write-host "" 
                write-host "Please provide the Report(s) you want to display:   " -NoNewline
                $choice = Read-Host
                Write-Host ""

                $ok = $choice -match '^[abcdefghijklmnopx]+$'

                if ( -not $ok) { write-host "Invalid selection" }
            }until ($ok)

            switch -Regex ( $choice ) {
                "A" {$arrfunctions += @(get-item function:GetDomainInfo),
                                        (get-item function:GetLicenses),
                                        (get-item function:O365RolesAndMembers),
                                        (get-item function:GetAllMSOLUsers),
                                        (get-item function:GetUsersAndGroups),
                                        (get-item function:GetUsersAndLicenses),
                                        (get-item function:GetAllDevices),
                                        (get-item function:GetMailboxes),
                                        (get-item function:GetMailcontacts),
                                        (get-item function:GetSPODetails),
                                        (get-item function:GroupsReport),
                                        (get-item function:GetSkypeForBusiness),
                                        (get-item function:SkypeForBUsage),
                                        (get-item function:skypeForBUserActivities),
                                        (get-item function:SFBClientDeviceDetailedReport)
                    }
                "B" {$arrfunctions += @(get-item function:GetDomainInfo)}
                "C" {$arrfunctions += @(get-item function:GetLicenses)}
                "D" {$arrfunctions += @(get-item function:O365RolesAndMembers)}
                "E" {$arrfunctions += @(get-item function:GetAllMSOLUsers)}
                "F" {$arrfunctions += @(get-item function:GetUsersAndGroups)}
                "G" {$arrfunctions += @(get-item function:GetUsersAndLicenses)}
                "H" {$arrfunctions += @(get-item function:GetAllDevices)}
                "I" {$arrfunctions += @(get-item function:GetMailboxes)}
                "J" {$arrfunctions += @(get-item function:GetMailcontacts)}
                "K" {$arrfunctions += @(get-item function:GetSPODetails)}
                "L" {$arrfunctions += @(get-item function:GroupsReport)}
                "M" {$arrfunctions += @(get-item function:GetSkypeForBusiness)}
                "N" {$arrfunctions += @(get-item function:SkypeForBUsage)}
                "O" {$arrfunctions += @(get-item function:skypeForBUserActivities)}
                "P" {$arrfunctions += @(get-item function:SFBClientDeviceDetailedReport)}
                "X" {exit} 
            }   
            $arrfunctions += @((get-item function:CreateExcelReportFile))
            return $arrfunctions
        } 
        
        Catch{}
    }
    End{}
}

Function OpenReportFile{
    Param(
        [Parameter(Mandatory=$True)]$LicenseReportFile
    )
  
    Begin{logwrite -Logstring "The report file is being opened" -type Informational}
  
    Process{
        Try{
            [threading.thread]::CurrentThread.CurrentCulture = 'en-US'
            $Excel = New-Object -ComObject excel.application 
		    $Excel.workbooks.open($LicenseReportFile) |out-null
		    $Excel.Visible = $True
        } 
        
        Catch{
            logwrite -Logstring $error -type Error;Break
        }
    }
  
    End{logwrite -Logstring "Successfully opened the Reportfile." -type Informational }
}

Function GetDomainInfo{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get domain information of connected domains" -type Informational;$i=1}
  
    Process{
        Try{
            $datastring = @()
            $domains = Get-MsolDomain
            $totaldomains = $domains.count
            Foreach($domain in $domains){
                write-progress -ParentId 1 -id 2 -activity "Get domain information of connected domains" -status 'Running->' -percentcomplete ($i/$totaldomains*100) -currentOperation "Domain found: $($domain.name)"
			    $hash = @{'DomainName' = $domain.name;'Capabilities' = $domain.capabilities;'Authentication' = $domain.Authentication;'IsDefault' = $domain.IsDefault;'Status' = $domain.Status;'IsInitial' = $domain.IsInitial}
			    $datastring += New-Object PSObject -Property $hash
                $hash = $null
                $i++
		    }
            $datastring | Select DomainName, Capabilities, Authentication, IsDefault, Status,  IsInitial | Export-Csv -Path $DomainsCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
        write-progress -ParentId 1 -id 2 -activity "Get domain information of connected domains" -Completed
        logwrite -Logstring "Successfull added domain information to the logfile $($DomainsCsv)" -type Informational
         
    }
}

Function GetAllMSOLUsers{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get detailed information of all MSOL Users" -type Informational;$i=1}
  
    Process{
        Try{
            $datastring = @()
            $totalMSOLUsers =  $MSOLUsers.count
            Foreach($user in $MSOLUsers){
                write-progress -ParentId 1 -id 3 -activity "Getting user information" -status 'Running->' -percentcomplete ($i/$totalMSOLUsers*100) -currentOperation "User found: $($user.DisplayName)"
                if($user.LastDirSyncTime){$lastDirSyncTime = $user.LastDirSyncTime} else {$lastDirSyncTime = "CloudUser"}
	            $hash = @{'Title' = $user.title;'DisplayName' = $user.DisplayName;'Address' = $User.streetaddress;'PostalCode' = $user.postalcode;'City' = $user.City;'State' = $user.state ;'Phone' = $user.PhoneNumber;'MobilePhone' = $User.MobilePhone;'UsageLocation' = $user.UsageLocation;'Department' = $user.Department;'UserPrincipalName' = $user.userprincipalname;'SignInName' = $user.SignInName;'LastDirSyncTime' = $lastDirSyncTime;'WhenCreated' = $user.whencreated;'PasswordNeverExpires' = $user.PasswordNeverExpires;'LastPasswordChangeTimestamp' = $user.LastPasswordChangeTimestamp}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
                $user=$null
                $i++
		    }
            $datastring |Select Title,DisplayName,Address,PostalCode,City,State,Phone,MobilePhone,UsageLocation,Department,UserPrincipalName,SignInName,PasswordNeverExpires,LastDirSyncTime,WhenCreated,LastPasswordChangeTimestamp |Export-Csv -Path $AllMSOLUsers -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
        write-progress -ParentId 1 -id 3 -activity "Getting user information" -Completed
        logwrite -Logstring "Successfull added detailed information of all MSOL Users to the logfile $($AllMSOLUsers)" -type Informational 
    }
}

Function GetAllDevices{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get used devices by MSOL Users" -type Informational;$i=1}
  
    Process{
        Try{
            $datastring = @()
            $totalMSOLUsers =  $MSOLUsers.count
            Foreach($user in $MSOLUsers){
                write-progress -ParentId 1 -id 4 -activity "Getting devices of all users" -status 'Running->' -percentcomplete ($i/$totalMSOLUsers*100) -currentOperation "User found: $($user.DisplayName)"
                $userDevices = Get-MsolDevice -RegisteredOwnerUpn $user.userprincipalname
                if ($Userdevices){
                    $totaldevices = $Userdevices.count
                    $j = 1
                    Foreach($device in $Userdevices){
                        write-progress -id 5 -ParentId 4 -activity "Getting device" -status 'Running->' -percentcomplete ($j/$totaldevices*100) -currentOperation "Device found: $($device.displayname)"
	                    $hash = @{'DisplayName' = $user.DisplayName;'UserPrincipalName' = $user.userprincipalname;'DeviceName' = $device.displayname;'DeviceID' = $device.deviceID;'OS_Type' = $device.DeviceOSType;'OS_Version' =$device.deviceOSVersion;'TrustType' = $device.DeviceTrustType;'TrustLevel' = $Device.DeviceTrustLevel;'LastLogonStamp' = $Device.ApproximateLastLogonTimestamp;'RegisteredOwner' = $Device.RegisteredOwners}
                        $datastring += New-Object PSObject -Property $hash
                        $hash = $null
                        $j++
                    }
                    $Userdevices = $null
                }
                $i++
                write-progress -id 5 -ParentId 4 -activity "Getting device" -Completed
		    }
            $datastring |Select DisplayName,UserPrincipalName,DeviceName,DeviceID,OS_Type,OS_Version,TrustType,TrustLevel,LastLogonStamp,RegisteredOwner |Export-Csv -Path $DevicesCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
        write-progress -ParentId 1 -id 4 -activity "Getting devices of all users" -Completed
        logwrite -Logstring "Successfull added devices used by MSOL Users to the logfile $($AllMSOLUsers)" -type Informational
    }

}

Function GetSPODetails{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get detailed information of the SharePoint Online Sites" -type Informational;$i=1}
  
    Process{
        Try{
            $datastring = @()
            $SPOSites = Get-SPOSite -Detailed
            $TotalSites = $SPOSites.count
            if($TotalSites -eq $null){$TotalSites = 1}
            Foreach($site in $SPOSites){
                write-progress -ParentId 1 -id 6 -activity "Getting SPO sites" -status 'Running->' -percentcomplete ($i/$totalSites*100) -currentOperation "SPO site found: $($site.title)"
                $owner = $null
                if($site.owner -notlike "S-1-5-*"){
                    $owner = $site.owner
                } 
	            $hash = @{'Title' = $site.title;'Owner' = $owner;'URL' = $site.url;'CurrentStorageUsage' = $site.StorageUsageCurrent;'SharingCapability' = $site.SharingCapability;'StorageQuota' = $site.StorageQuota;'StorageQuotaWarningLevel' = $site.StorageQuotaWarningLevel;'Template' = $site.template;'LocaleId' = $site.LocaleId;'NumberOfSubSites' = $site.WebsCount}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
                $i++
		    }
            $datastring |Select Title,URL,Owner,CurrentStorageUsage,NumberOfSubSites,SharingCapability,Template,StorageQuota,StorageQuotaWarningLevel,LocaleId |Export-Csv -Path $SPOTenantCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
         write-progress -ParentId 1 -id 6 -activity "Getting SPO sites" -Completed
        logwrite -Logstring "Successfull added SharePoint Online Site information to the logfile $($SPOTenantCsv)" -type Informational
    }
}

Function GetMailboxes{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get information of the Mailboxes" -type Informational;$i=1}
  
    Process{
        Try{
            $datastring = @()
            $Mailboxes = Get-Mailbox -ResultSize Unlimited
            $TotalMailboxes = $Mailboxes.count
            if($TotalMailboxes -eq $null){$TotalMailboxes = 1}
            Foreach($Mailbox in $Mailboxes){
                write-progress -ParentId 1 -id 7 -activity "Getting mailboxes" -status 'Running->' -percentcomplete ($i/$TotalMailboxes*100) -currentOperation "Mailbox found: $($mailbox.primarysmtpaddress)"
                $MailboxStatistics = Get-MailboxStatistics -Identity $mailbox.primarysmtpaddress
                $MailboxPermissionsSendAs = Get-RecipientPermission  -Identity $Mailbox.id | Where { ($_.IsInherited -eq $False) -and -not ($_.Trustee -like “NT AUTHORITY\SELF”) } | Select Trustee
                $SendOnBehalfOfUsers = ($MailboxPermissionsSendAs.trustee -join ",")
                $MailboxPermissionsFullAccess = Get-MailboxPermission -identity $Mailbox.id | Where-Object {$_.User -ne "NT AUTHORITY\SELF" -and $_.IsInherited -ne $true -and $_.user -notlike "S-1-5-*" -and $_.AccessRights -contains "FullAccess"}
                $FullAccessUsers = ($MailboxPermissionsFullAccess.user -join ",")
	            $hash = @{'PrimarySMTP' = $mailbox.primarysmtpaddress;'OtherMailAddresses' = $mailbox.emailaddresses;'AuditEnabled' = $mailbox.AuditEnabled;'RetainDeletedItemsFor' = $mailbox.RetainDeletedItemsFor;'Languages' = $mailbox.languages;'GrandSendOnBehalfto' = $SendOnBehalfOfUsers;'FullAccessRights' = $FullAccessUsers;'RecipientTypeDetails' = $mailbox.RecipientTypeDetails;'ResourceType' = $mailbox.ResourceType;'Capacity'= $Mailbox.ResourceCapacity; 'Location' = $Mailbox.Office;'LitigationHoldEnabled' = $mailbox.LitigationHoldEnabled ;'ArchiveStatus' = $mailbox.ArchiveStatus;'DisplayName' = $MailboxStatistics.DisplayName;'TotalItems'=$MailboxStatistics.ItemCount;'TotalSize'=$MailboxStatistics.TotalItemSize;'DeletedItemSize'=$MailboxStatistics.TotalDeletedItemSize}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
                $i++
		    }
            $datastring |Select DisplayName,PrimarySMTP,RecipientTypeDetails,ResourceType,Capacity,Location,RetainDeletedItemsFor,Languages,AuditEnabled,GrandSendOnBehalfto,FullAccessRights,LitigationHoldEnabled,ArchiveStatus,TotalItems,TotalSize,DeletedItemSize,OtherMailAddresses|Export-Csv -Path $mailboxCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
         write-progress -ParentId 1 -id 7 -activity "Getting usermailboxes" -Completed
        logwrite -Logstring "Successfull added User Mailbox information to the logfile $($UsermailboxCsv)" -type Informational }
}

Function GetMailcontacts{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get information of the mail contacts" -type Informational;$i=1}
  
    Process{
        Try{
            $datastring = @()
            $contacts = Get-mailcontact -ResultSize Unlimited
            $TotalContacts = $Contacts.count
            if($TotalContacts -eq $null){$TotalContacts = 1}
            Foreach($contact in $contacts){
                write-progress -ParentId 1 -id 8 -activity "Getting all contacts" -status 'Running->' -percentcomplete ($i/$TotalContacts*100) -currentOperation "Contact found: $($contact.identity)"
                $hash = @{'ContactName' = $contact.identity;'E-Mail' = $contact.PrimarySmtpAddress;'HiddenFromAddresslist' = $contact.HiddenFromAddressListsEnabled}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
                $i++
		    }
            $datastring |Select ContactName,E-Mail,HiddenFromAddresslist |Export-Csv -Path $MailContactCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
        write-progress -ParentId 1 -id 8 -activity "Getting all contacts" -Completed
        logwrite -Logstring "Successfull added mail contacts to the logfile $($MailContactCsv)" -type Informational }
}

Function GetUsersAndGroups{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get information of Users and Groups";$i=1}
  
    Process{
        Try{
            $datastring = @()
            $SecurityGroups = Get-MsolGroup -All
            $totalSecurityGroups = $SecurityGroups.count
            if($totalSecurityGroups -eq $null){$totalSecurityGroups = 1}
            Foreach($group in $SecurityGroups){
                write-progress -ParentId 1 -id 9 -activity "Getting all Security groups" -status 'Running->' -percentcomplete ($i/$totalSecurityGroups*100) -currentOperation "Security group found: $($group.DisplayName)"
                $GroupMembers = get-msolgroupMember -GroupObjectId $Group.objectID -All
                $totalGroupMembers = $GroupMembers.count
                $j=1
                Foreach($member in $GroupMembers){
                    write-progress -ParentId 9 -id 10 -activity "Getting members of $group.DisplayName" -status 'Running->' -percentcomplete ($j/$totalGroupMembers*100) -currentOperation "Groupmember found: $($Member.DisplayName)"
                    #It is possible to get more information of the group members, such as Title etc. but be aware the script will take minimal twice the normal time it need to run!
                    #If($Member.EmailAddress){$MemberTitle = $(get-msoluser -UserPrincipalName $Member.EmailAddress |select Title).title}
                    #$hash = @{'GroupName' = $group.DisplayName;'Member_DisplayName' = $Member.DisplayName;'Member_EMailAddress' = $Member.EmailAddress;'Title' = $MemberTitle }
                    $hash = @{'GroupName' = $group.DisplayName;'EmailAddress' = $group.EmailAddress;'Member_DisplayName' = $Member.DisplayName;'Member_EMailAddress' = $Member.EmailAddress}
                    $datastring += New-Object PSObject -Property $hash
                    $hash = $null
                    $j++
                }
                write-progress -ParentId 9 -id 10 -activity "Getting members of $group.DisplayName" -Completed
                $i++
		    }
            $datastring |Select GroupName,EmailAddress,Member_DisplayName,Member_EMailAddress|Export-Csv -Path $ADGroupsCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
       write-progress -ParentId 1 -id 9 -activity "Getting all Security groups" -Completed
        logwrite -Logstring "Successfull added User and Group information to the logfile $($ADGroupsCsv)" -type Informational }
}

Function GetUsersAndLicenses{
    Param(
        
    )
  
    Begin{
        logwrite -Logstring "Get license information of Users"
        #Create a look up table for Human friendly license names
        $SKUs = @{
                "AAD_BASIC"                     = "Azure Active Directory Basic";`
                "AAD_PREMIUM"                   = "Azure Active Directory Premium";`
                "ADALLOM_STANDALONE"            = "Cloud App Security";`
                "ATP_ENTERPRISE"                = "Exchange Online ATP";`
                "BI_AZURE_P1"                   = "Power BI Reporting and Analytics";`
                "CRMIUR"                        = "Dynamics CRM Online Pro IUR"
                "CRMPLAN1"                      = "Dynamics CRM Online Essential";`
                "CRMPLAN2"                      = "Dynamics CRM Online Basic" ;`
                "CRMSTANDARD"                   = "Dynamics CRM Online Pro";`
                "DESKLESSPACK"                  = "O365 Enterprise K1";`
                "DESKLESSPACK_YAMMER"           = "Office 365 Enterprise K1 With Yammer";`
                "EMS"                           = "Enterprise Mobility And Security Suite";`
                "ENTERPRISEPACK"                = "O365 Enterprise E3";`
                "ENTERPRISEPREMIUM"             = "O365 Enterprise E5";`
                "ENTERPRISEPREMIUM_NOPSTNCONF"  = "O365 Enterprise E5 w/o PSTN Conf";`
                "ENTERPRISEWITHSCAL"            = "O365 Enterprise E4";`
                "EOP_ENTERPRISE"                = "Exchange Online Protection";`
                "EQUIVIO_ANALYTICS"             = "O365 Advanced eDiscovery";`
                "ERP_INSTANCE"                  = "Microsoft Power BI for Office 365";`
                "EXCHANGEARCHIVE"               = "EOA for Exchange Server";`
                "EXCHANGEARCHIVE_ADDON"         = "EOA for Exchange Online";`
                "EXCHANGEDESKLESS"              = "Exchange Online Kiosk";`
                "EXCHANGEENTERPRISE"            = "Exchange Online (Plan 2)";`
                "EXCHANGESTANDARD"              = "Exchange Online (Plan 1)";`
                "EXCHANGE_ANALYTICS"            = "Delve Analytics";`
                "INTUNE_A"                      = "Intune";`
                "INTUNE_STORAGE"                = "Intune Extra Storage";`
                "LITEPACK"                      = "O365 Small Business";`
                "LITEPACK_P2"                   = "O365 Small Business Premium";`
                "LOCKBOX"                       = "Customer Lockbox";`
                "MCOEV"                         = "SfB Cloud PBX";`
                "MCOIMP"                        = "SfB Online (Plan 1)";`
                "MCOMEETADV"                    = "SfB PSTN Conferencing";`
                "MCOPLUSCAL"                    = "SfB Plus CAL";`
                "MCOPSTN1"                      = "SfB PSTN Dom. Calling";`
                "MCOPSTN2"                      = "SfB PSTN Dom. and Int. Calling";`
                "MCOSTANDARD"                   = "SfB Online (Plan 2)";`
                "O365_BUSINESS"                 = "O365 Business";`
                "O365_BUSINESS_ESSENTIALS"      = "O365 Business Essentials";`
                "O365_BUSINESS_PREMIUM"         = "O365 Business Premium";`
                "OFFICESUBSCRIPTION"            = "O365 ProPlus";`
                "PLANNERSTANDALONE"             = "Office 365 Planner";`
                "POWERAPPS_INDIVIDUAL_USER"     = "Microsoft PowerApps and logical Streams";`
                "POWER_BI_ADDON"                = "Power BI Add-on";`
                "POWER_BI_INDIVIDUAL_USE"       = "Power BI Individual User";`
                "POWER_BI_PRO"                  = "Power BI (Pro)";`
                "POWER_BI_STANDALONE"           = "Power BI Stand Alone";`
                "POWER_BI_STANDARD"             = "Power BI (free)";`
                "PROJECTCLIENT"                 = "Project Pro for O365";`
                "PROJECTESSENTIALS"             = "Project Lite";`
                "PROJECTONLINE_PLAN_1"          = "Project Online";`
                "PROJECTONLINE_PLAN_2"          = "Project Online and Pro";`
                "RIGHTSMANAGEMENT"              = "Azure Rights Management Premium";`
                "RIGHTSMANAGEMENT_ADHOC"        = "Windows Azure Rights Management";`
                "SHAREPOINTENTERPRISE"          = "SharePoint Online (Plan 2)";`
                "SHAREPOINTSTANDARD"            = "SharePoint Online (Plan 1)";`
                "SHAREPOINTSTORAGE"             = "O365 Extra File Storage";`
                "STANDARDPACK"                  = "O365 Enterprise E1";`
                "STANDARDWOFFPACK"              = "O365 Enterprise E2 (Nonprofit E1)";`
                "STREAM"                        = "Microsoft Stream"; `
                "VISIOCLIENT"                   = "Visio Pro for O365";`
                "WACONEDRIVEENTERPRISE"         = "OneDrive for Business (Plan 2)";`
                "WACONEDRIVESTANDARD"           = "OneDrive for Business (Plan 1)";`
                "YAMMER_ENTERPRISE_STANDALONE"  = "Yammer Enterprise"
        }
        $i=1

    }
  
    Process{
        Try{
            $datastring = @()
            $totalusers = $MSOLUsers.count
            Foreach($user in $MSOLUsers){
                write-progress -ParentId 1 -id 10 -activity "Finding users" -status 'Running->' -percentcomplete ($i/$totalusers*100) -currentOperation "User found: $($user.DisplayName)"
                if($user.LastDirSyncTime -eq $null){$Synched = $false} else{$Synched = $True}
                foreach($sku in $user.licenses){
                    if($skus.ContainsKey($sku.AccountSkuID.substring($tenant.length+1).ToString())){
                        $lic = $SKUs.get_item($sku.AccountSkuID.substring($tenant.length+1)).ToString()
                        $hash = @{'DisplayName' = $user.DisplayName;'FirstName' = $user.FirstName;'LastName' = $user.Lastname;'UPN' = $user.UserPrincipalName;'Synched' = $Synched;'Islicensed' = $user.IsLicensed;'License' = $lic}
                        $datastring += New-Object PSObject -Property $hash
                        $lic = $null
                        $hash = $null
                    }
                    else{
                        logwrite -Logstring "The value $($sku.AccountSkuID.substring($tenant.length+1)) does not exists in the hash table, please add it to the hashtable and provide a user friendly name for it!" -type Error
                    }
                }
                $Synched = $null
		    }
            $i++
            $datastring |Select DisplayName,FirstName,LastName,UPN,Synched,IsLicensed,License|Export-Csv -Path $LicensesCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{
        write-progress -ParentId 1 -id 10 -activity "Finding users" -Completed
        logwrite -Logstring "Successfull added User and license information to the logfile $($LicensesCsv)" -type Informational }
}

Function GetLicenses{
    Param(
        
    )
  #Get-MsolSubscription | fl
    Begin{logwrite -Logstring "Add license information of the organization"}
  
    Process{
        Try{
            $datastring = @()
            $accountSKU = Get-MsolAccountSku
            foreach($sku in $accountSKU){
                If($sku.SubscriptionIds.count -gt 1){
                    foreach($SubscriptionId in $sku.SubscriptionIds){ 
                        $SubLics = Get-MsolSubscription -SubscriptionId $($SubscriptionId).Guid
                        $hash = @{'AccountSkuID' = $sku.AccountSkuId;'ActiveUnits' = $sku.ActiveUnits;'ConsumedUnits' = $sku.ConsumedUnits;'LockedoutUnits' = $sku.LockedoutUnits;'DateCreated' = $SubLics.DateCreated;'NextLifeCycleDate' = $SubLics.NextLifecycleDate;'IsTrail' = $SubLics.IsTrial;'Status'=$SubLics.Status;'SubscriptionID GUID' = $($SubscriptionId).Guid}
                        $datastring += New-Object PSObject -Property $hash
                    }
                }
                else{
                    $SubLics = Get-MsolSubscription -SubscriptionId $($sku.SubscriptionIds).Guid
                    $hash = @{'AccountSkuID' = $sku.AccountSkuId;'ActiveUnits' = $sku.ActiveUnits;'ConsumedUnits' = $sku.ConsumedUnits;'LockedoutUnits' = $sku.LockedoutUnits;'DateCreated' = $SubLics.DateCreated;'NextLifeCycleDate' = $SubLics.NextLifecycleDate;'IsTrail' = $SubLics.IsTrial;'Status'=$SubLics.Status}
                    $datastring += New-Object PSObject -Property $hash
                }
            }
            $datastring |Select AccountSkuID,ActiveUnits,ConsumedUnits,LockedoutUnits,DateCreated,NextLifeCycleDate,IsTrial,Status|Export-Csv -Path $LicenseCountCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 
        
        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfull added license information of the organization to the logfile $($LicenseCountCsv)" -type Informational }
}

Function GetSkypeForBusiness{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get Skype for Business information"}
  
    Process{
        Try{
            $datastring = @()
            $AllowedDomains = $(Get-CsTenantFederationConfiguration | Select-Object -ExpandProperty AllowedDomains | Select-Object AllowedDomains).AllowedDomains
            $BlockedDomains = $(Get-CsTenantFederationConfiguration | Select-Object -ExpandProperty BlockedDomains | Select-Object domain).domain
            $AllowPublicUsers = $(Get-CsTenantFederationConfiguration |Select-Object AllowPublicUsers | Select-Object AllowPublicUsers).AllowPublicUsers
            If ($AllowedDomains.count -eq 0){$AllowedDomains = "All"}
            If ($BlockedDomains.count -eq 0){$BlockedDomains = "None"}
            $hash = @{'AllowedDomains' = $AllowedDomains;'BlockedDomains' = $BlockedDomains;'AllowPublicUsers' = $AllowPublicUsers}
            $datastring += New-Object PSObject -Property $hash
            $datastring |Select AllowedDomains,BlockedDomains,AllowPublicUsers |Export-Csv -Path $SkypeForBCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 

        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfull added Skype for Business information to the logfile $($SkypeForBCsv)" -type Informational }
}

Function O365RolesAndMembers{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get Office 365 roles and members"}
  
    Process{
        Try{
            $datastring = @()
            $O365roles = Get-MsolRole 
            foreach($role in $O365roles){
                $RoleMembers = Get-MsolRoleMember -RoleObjectId $role.ObjectId |select EmailAddress,DisplayName,IsLicensed
                    foreach($roleMember in $RoleMembers){
                        $hash = @{'RoleName' = $role.Name;'DisplayName' = $roleMember.displayname;'Emailaddress' = $roleMember.emailaddress;'IsLicensed' = $role.isLicensed}
                        $datastring += New-Object PSObject -Property $hash
                    }
                    $hash = $null
            }
            $datastring |Select RoleName,DisplayName,Emailaddress,IsLicensed |Export-Csv -Path $O365RolesAndMembers -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 

        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfull added Office 365 roles and members to the logfile $($O365RolesAndMembers)" -type Informational }
}

function skypeForBUsage{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get Skype for Business Usage"}
  
    Process{
        Try{
            $datastring = @()
            $ActiveUserReport = Get-CsActiveUserReport -ReportType Daily -ResultSize unlimited
            foreach($activity in $ActiveUserReport){
                $hash = @{'Date' = $activity.Date;'ActiveUsers' = $activity.ActiveUsers;'ActiveIMUsers' = $activity.ActiveIMUsers;'ActiveAudioUsers' =  $activity.ActiveAudioUsers;'ActiveVideoUsers' =  $activity.ActiveVideoUsers;'ActiveApplicationSharingUsers' = $activity.ActiveApplicationSharingUsers; 'ActiveFileTransferUsers' = $activity.ActiveFileTransferUsers;'ActivePSTNConferencingUsers' = $activity.ActivePSTNConferencingUsers}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
            }
            $datastring |Select Date,ActiveUsers,ActiveIMUsers,ActiveAudioUsers,ActiveVideoUsers,ActiveApplicationSharingUsers,ActiveFileTransferUsers,ActivePSTNConferencingUsers |Export-Csv -Path $skypeForBUsageCSV -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 

        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfull added Skype for Business usage to the log file: $($SkypeForBUsageCsv)" -type Informational }
    
}

function SkypeForBUserActivities{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get Skype for Business User Activities"}
  
    Process{
        Try{
            $datastring = @()
            $ActiveUserActivityReport = Get-CsUserActivitiesReport -ResultSize Unlimited 
            foreach($Useractivity in $ActiveUserActivityReport){
                $hash = @{'TenantName' = $Useractivity.Tenantname;'Date' = $Useractivity.Date; 'UserName' =  $Useractivity.UserName; 'LastLogonTime' = $Useractivity.LastLogonTime; 'LastActivityTime' = $Useractivity.LastActivityTime; 'TotalP2PSessions' = $Useractivity.TotalP2PSessions; 'TotalP2PIMSessions' = $Useractivity.TotalP2PIMSessions; 'TotalP2PAudioSessions' = $Useractivity.TotalP2PAudioSessions; 'TotalP2PVideoSessions' = $Useractivity.TotalP2PVideoSessions;'TotalP2PApplicationSharingSessions' = $Useractivity.TotalP2PApplicationSharingSessions; 'TotalP2PAudioSessionMinutes' = $Useractivity.TotalP2PAudioSessionMinutes; 'TotalP2PVideoSessionMinutes' =$Useractivity.TotalP2PVideoSessionMinutes; 'TotalOrganizedConferences' =  $Useractivity.TotalOrganizedConferences; 'TotalOrganizedIMConferences' = $Useractivity.TotalOrganizedIMConferences; 'TotalOrganizedAVConferences' =  $Useractivity.TotalOrganizedAVConferences; 'TotalOrganizedApplicationSharingConferences' = $Useractivity.TotalOrganizedApplicationSharingConferences; 'TotalOrganizedWebConferences' = $Useractivity.TotalOrganizedWebConferences; 'TotalOrganizedDialInConferences' =  $Useractivity.TotalOrganizedDialInConferences; 'TotalOrganizedAVConferenceMinutes'  = $Useractivity.TotalOrganizedAVConferenceMinutes; 'TotalParticipatedConferences' =   $Useractivity.TotalParticipatedConferences; 'TotalParticipatedIMConferences' =  $Useractivity.TotalParticipatedIMConferences; 'TotalParticipatedAVConferences' =  $Useractivity.TotalParticipatedAVConferences; 'TotalParticipatedApplicationSharingConferences' = $Useractivity.TotalParticipatedApplicationSharingConferences; 'TotalParticipatedWebConferences' =  $Useractivity.TotalParticipatedWebConferences; 'TotalParticipatedDialInConferences' = $Useractivity.TotalParticipatedDialinConferences; 'TotalParticipatedAVConferenceMinutes' = $Useractivity.TotalParticipatedAVConferenceMinutes; 'TotalPlacedPSTNCalls' = $Useractivity.TotalPlacedPSTNCalls; 'TotalReceivedPSTNCalls' = $Useractivity.TotalReceivedPSTNCalls; 'TotalPlacedPSTNCallMinutes' = $Useractivity.TotalPlacedPSTNCallMinutes; 'TotalReceivedPSTNCallMinutes' = $Useractivity.TotalReceivedPSTNCallMinutes; 'TotalMessages' = $Useractivity.TotalMessages; 'TotalTransferedFiles' =  $Useractivity.TotalTransferedFiles}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
            }
            $datastring |Select TenantName,Date,UserName,LastLogonTime,LastActivityTime,TotalP2PSessions,TotalP2PIMSessions,TotalP2PAudioSessions,TotalP2PVideoSessions,TotalP2PApplicationSharingSessions,TotalP2PAudioSessionMinutes,TotalP2PVideoSessionMinutes,TotalOrganizedConferences,TotalOrganizedIMConferences,TotalOrganizedAVConferences,TotalOrganizedApplicationSharingConferences,TotalOrganizedWebConferences,TotalOrganizedDialInConferences,TotalOrganizedAVConferenceMinutes,TotalParticipatedConferences,TotalParticipatedIMConferences,TotalParticipatedAVConferences,TotalParticipatedApplicationSharingConferences,TotalParticipatedWebConferences,TotalParticipatedDialInConferences,TotalParticipatedAVConferenceMinutes,TotalPlacedPSTNCalls,TotalReceivedPSTNCalls,TotalPlacedPSTNCallMinutes,TotalReceivedPSTNCallMinutes,TotalMessages,TotalTransferedFiles |Export-Csv -Path $skypeForBUserActivitiesCSV -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 

        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfully added Skype for Business User Activities to the log file: $($SkypeForBUsageCsv)" -type Informational }
    
}

function SFBClientDeviceDetailedReport{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get Skype for Business Client Device Detailed Report"}
  
    Process{
        Try{
            $datastring = @()
            $ClientDeviceDetailedReport =Get-CsClientDeviceDetailReport -ResultSize Unlimited 
            foreach($Device in $ClientDeviceDetailedReport){
                $hash = @{'TenantName' = $Device.Tenantname;'Date' = $Device.Date; 'UserName' =  $Device.UserName; 'Windows' = $Device.WindowsActivities; 'WindowsPhone' = $Device.WindowsPhoneActivities; 'Android' = $Device.AndroidActivities; 'iPhone' = $Device.iPhoneActivities; 'iPad' = $Device.iPadActivities}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
            }
            $datastring |Select TenantName,Date,UserName,Windows,WindowsPhone,Android,iPhone,iPad |Export-Csv -Path $SFBClientDeviceDetailedReportCSV -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 

        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfully added Skype for BusinessClient Device Detailed report to the log file: $($SFBClientDeviceDetailedReportCSV)" -type Informational }
    
}

function GroupsReport{
    Param(
        
    )
  
    Begin{logwrite -Logstring "Get Office 365 Groups Report"}
  
    Process{
        Try{
            $datastring = @()
            $O365Groups = Get-UnifiedGroup -ResultSize Unlimited 
            foreach($group in $O365Groups){
                $ManagedBy = $group.ManagedByDetails -join ","
                $hash = @{'Alias' = $group.Alias;'Name' = $group.Name;'ManagedBy'=$ManagedBy;'AccessType' = $group.AccessType;'AutoSubscribeNewMembers' = $group.AutoSubscribeNewMembers;'WelcomeMessageEnabled'=$group.WelcomeMessageEnabled;'AllowAddGuests' = $group.GroupType;'PrimarySmtpAddress' = $group.PrimarySmtpAddress;'SharePointSiteUrl' = $group.SharePointSiteUrl;'SharePointDocumentsUrl' = $group.SharePointDocumentsUrl;'SharePointNotebookUrl'=$group.SharePointNotebookUrl;'WhenCreated' = $group.WhenCreated;'Notes'=$group.Notes}
                $datastring += New-Object PSObject -Property $hash
                $hash = $null
            }
            $datastring |Select Alias,Name,ManagedBy,AccessType,AutoSubscribeNewMembers,WelcomeMessageEnabled,AllowAddGuests,PrimarySmtpAddress,SharePointSiteUrl,SharePointDocumentsUrl,SharePointNotebookUrl,WhenCreated,Notes |Export-Csv -Path $GroupsCsv -Encoding UTF8 -NoTypeInformation -Delimiter ","
        } 

        Catch{logwrite -Logstring $error -type Error;Break}
    }
  
    End{logwrite -Logstring "Successfully added Skype the Groups Detailed report to the log file: $($GroupsCsv)" -type Informational }
    
}

Function CreateExcelReportFile{
    Param(
        
    )
  
    Begin{
        logwrite -Logstring "Creating the Excel Report File"
        $xlDatabase 			= 1
	    $xlPivotTableVersion12 	= 3
	    $xlHidden              	= 0
	    $xlRowField            	= 1
	    $xlColumnField        	= 2
	    $xlPageField           	= 3
	    $xlDataField           	= 4
   }
  
    Process{
        Try{
            [threading.thread]::CurrentThread.CurrentCulture = 'en-US'
            $Excel = New-Object -ComObject excel.application
        } 
        Catch{logwrite -Logstring $error -type Error;Break}
        
        $Excel.visible = $false
        $Excel.DisplayAlerts = $false
        $workbooks = $excel.Workbooks.Add()
        $worksheets = $workbooks.worksheets
        Try{
            #for ($i=1; $i -le ($csvReports.count-1); $i++){$worksheets.add()|Out-Null}
        }
        Catch{
            logwrite -Logstring $error -type Error;Break
            Worksheets
        }
        $i=1
        #Create the sheets and import the csv file
        ForEach($csvReport in $csvReports.GetEnumerator()|Sort-Object Name -descending){
            $worksheets.add() |out-null
            $worksheets.item($i).name = $csvReport.name
            if($csvReport.name -ne "UserLicenses"){
                $TxtConnector = ("TEXT;" + $csvReport.value)
                $cellref = $worksheets.item($csvReport.name).Range("A1")
                $Connector = $worksheets.item($csvReport.name).QueryTables.add($TxtConnector,$CellRef)
                $worksheets.item($csvReport.name).QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
                $worksheets.item($csvReport.name).QueryTables.item($Connector.name).TextFileParseType  = 1
                $worksheets.item($csvReport.name).QueryTables.item($Connector.name).Refresh()|Out-Null
                $worksheets.item($csvReport.name).QueryTables.item($Connector.name).Delete()|Out-Null
                $worksheets.item($csvReport.name).UsedRange.EntireColumn.AutoFit()|out-null
                $worksheets.item($csvReport.name).rows(1).Font.Bold = $true
            }
        }
       
        Foreach($sheet in $worksheets){
            $sheet.Activate()
            $sheet.Application.ActiveWindow.SplitRow = 1
            $sheet.Application.ActiveWindow.SplitColumn = 1
            $sheet.Application.ActiveWindow.FreezePanes = $true
            $Rng = $sheet.Range("A1").CurrentRegion
            $rng.AutoFilter()|Out-Null
            
            If($sheet.name -eq "UserAndGroups"){
                $usrNGrps = $worksheets.item("UserAndGroups")
                $Rng = $usrNGrps.Range("A1").CurrentRegion
                $rowCount = ($Rng.rows.count)
                $rng2 = $usrNGrps.Range("A2:D$rowCount")
                $rng2.Sort($rng2,1)|Out-Null
                $rng.Subtotal(1,-4112,(3),$true,$false,$true)|out-null
                $usrNGrps.Outline.ShowLevels(2)|Out-Null
                $usrNGrps.Activate()
            }


            If($sheet.name -eq "Devices"){
                $Devices = $worksheets.item("Devices")
                $Rng = $Devices.Range("A1").CurrentRegion
                $rowCount = ($Rng.rows.count)
                $rng2 = $Devices.Range("A2:D$rowCount")
                $rng2.Sort($rng2,1)|Out-Null
                $rng.Subtotal(1,-4112,(2),$true,$false,$true)|out-null
                $Devices.Outline.ShowLevels(2)|Out-Null
                $Devices.Activate()
            }

            If($sheet.name -eq "AllMSOLUsers"){
                $AllMSOLUsrs = $worksheets.item("AllMSOLUsers")
                $Rng = $AllMSOLUsrs.Range("A1").CurrentRegion
                $rowCount = ($Rng.rows.count)
                $rng2 = $AllMSOLUsrs.Range("A2:P$rowCount")
                $rng2.Sort($rng2,1)|Out-Null
                $rng.Subtotal(1,-4112,(2),$true,$false,$true)|out-null
                $AllMSOLUsrs.Outline.ShowLevels(2)|Out-Null
                $AllMSOLUsrs.Activate()
            }
            
            If($sheet.name -eq "Office365Roles"){
                $O365Roles = $worksheets.item("Office365Roles")
                $Rng = $O365Roles.Range("A1").CurrentRegion
                $rowCount = ($Rng.rows.count)
                $rng2 = $O365Roles.Range("A2:D$rowCount")
                $rng2.Sort($rng2,1)|Out-Null
                $rng.Subtotal(1,-4112,(2),$true,$false,$true)|out-null
                $O365Roles.Outline.ShowLevels(2)|Out-Null
                $O365Roles.Activate()
            }
            
            #create the Pivot Table for the UserLicenses report sheet
            If($sheet.name -eq "UserLicenses_RAW"){
                $Raw = $worksheets.item("UserLicenses_RAW")
                $Licenses = $worksheets.item("UserLicenses")
                $Rng = $Raw.Range("A1").CurrentRegion

                #create a string we can use to create the Pivot table
                [string]$range = $Raw.name + '!' + "R1C1:R" + ($rng.Cells.Rows.count) + "C"+($rng.Cells.Columns.count)

                #Create the PivotTable in the Licenses sheet
	            $PivotTable = $Workbooks.PivotCaches().Create($xlDatabase,$range,$xlPivotTableVersion12)
	            $PivotTable.CreatePivotTable($Licenses.range("A1"),"LicenceTable")|Out-Null
	            $PageField = $Licenses.PivotTables("LicenceTable").pivotfields("Synched")
	            $PageField.Orientation = $xlPageField
	            $RowFields = $Licenses.PivotTables("LicenceTable").pivotfields("DisplayName")
	            $RowFields.Orientation = $xlRowField
	            $columnField = $Licenses.PivotTables("LicenceTable").pivotfields("License")
	            $columnField.Orientation = $xlColumnField
	            $Datafields = $Licenses.PivotTables("LicenceTable").pivotfields("License")
	            $Datafields.Orientation = $xlDataField
	            $Licenses.UsedRange.EntireColumn.AutoFit()|Out-Null
                $Licenses.Activate()
                $Licenses.Application.ActiveWindow.SplitRow = 4
                $Licenses.Application.ActiveWindow.SplitColumn = 1
                $Licenses.Application.ActiveWindow.FreezePanes = $true

                #Hide the RAW data sheet, for future purposes (we could also delete it, since data is stored in the cache)
	            $Raw.Visible = $false
            }
        }
        
        $Worksheets.Item("Sheet1").Delete()
        $workbooks.saveas($ExcelReport)
	    $workbooks.Close($ExcelReport)
	    $Excel.Quit()
        
    }
  
    End{logwrite -Logstring "The Report file is created and save here: $($ExcelReport)" -type Informational }
}

#------------------------------------------------------[Login to MS Online Services]------------------------------------------------
$currentUser = "$env:userdomain\$env:username"
logwrite -logfile $Logfile -type Informational -Logstring "Script started on $(hostname) by $($currentUser)"
logwrite -logfile $Logfile -type Informational -Logstring "Used script version: $($ScriptVersion)"

#check if the script runs under Admin privileges
If($(Test-Administrator) -eq $false){
    logwrite -Logstring "The script cannot be run because it needs to be running as Administrator." -type Error
    logwrite -Logstring "The current Windows PowerShell session is not running as Administrator." -type Error
    logwrite -Logstring "Start Windows PowerShell by using the Run as Administrator option, and then try running the script again." -type Error
    Break
}

#Ask the user some questions
$Reportpath = OpenFileBrowserDialog "Please provide the location were the report file must be saved."
$OpenReport = DisplayMessageBox -Message "Do you want to open the report file after it has been generated?" -WindowTitle "Open Reportfile after creation"
$currentUser = "$env:userdomain\$env:username"

#Ask the user which reports need to be created
$reportsNeeded = whatToDo

#Login to the MSOLservices
$creds,$site = CreateCreds -site $site -user $User -password $Password
Login2MSOL -creds $creds
Login2SPO -creds $creds -site $site
Login2SfB -creds $creds
Login2ExO -creds $creds
$tenant = GetTenantID

#-----------------------------------------------------------ScriptActions------------------------------------------------------------
#Prepare the log files
#We first create a csv file and then we import this csv file in excel. This is much quicker then insert the data in Excel right away
$DomainsCsv                       = $logfilePath + "DomainsCsv - $tenant - $startTime.csv"
$AllMSOLUsers                     = $logfilePath + "AllMSOLUsers - $tenant - $startTime.csv"
$mailboxCsv                       = $logfilePath + "MailboxCsv - $tenant - $startTime.csv"
$MailContactCsv                   = $logfilePath + "MailContactCsv - $tenant - $startTime.csv"
$ADGroupsCsv                      = $logfilePath + "ADGroupCsv - $tenant - $startTime.csv"
$LicensesCsv                      = $logfilePath + "LicensesCsv - $tenant - $startTime.csv"
$LicenseCountCsv                  = $logfilePath + "LicenseCountCsv - $tenant - $startTime.csv"
$SPOTenantCsv                     = $logfilePath + "SPOTenantCsv - $tenant - $startTime.csv"
$SkypeForBCsv                     = $logfilePath + "SkypeForBCsv - $tenant - $startTime.csv"
$SkypeForBUsageCsv                = $logfilePath + "SkypeForBCsvUsage - $tenant - $startTime.csv"
$SkypeForBUserActivitiesCSV       = $logfilePath + "SkypeForBUserActivities - $tenant - $startTime.csv"
$SFBClientDeviceDetailedReportCSV = $logfilePath + "SFBClientDeviceDetailedReport - $tenant - $startTime.csv"
$DevicesCsv                       = $logfilePath + "DevicesCsv - $tenant - $startTime.csv"
$O365RolesAndMembers              = $logfilePath + "O365RolesAndMembers - $tenant - $startTime.csv"
$GroupsCsv                        = $logfilePath + "Groups - $tenant - $startTime.csv"

$ExcelReport = "$ReportPath\Office365 report - $tenant - $startTime.xlsx"

$csvReports = @{
    'Domains' = $DomainsCsv;
    'AvailableLicenses' = $LicenseCountCsv;
    'Office365Roles' = $O365RolesAndMembers;
    'AllMSOLUsers' = $AllMSOLUsers;
    'UserAndGroups' = $ADGroupsCsv;
    'UserLicenses_RAW' = $LicensesCsv;
    'UserLicenses' = "";
    'Devices' = $DevicesCsv;
    'Mailboxes' = $mailboxCsv;
    'MailContacts' = $MailContactCsv;
    'SharePointSites' = $SPOTenantCsv;
    'SkypeFB' = $SkypeForBCsv;
    'SkypeFBUsage' = $skypeForBUsageCsv;
    'SkypeFBActivities' = $SkypeForBUserActivitiesCSV;
    'SkypeFBDevices' = $SFBClientDeviceDetailedReportCSV;
    'O365Groups' = $GroupsCsv
}     

$MSOLUsers = Get-MsolUser -all

$TotalFunctions = $reportsNeeded.count
$i = 1
$reportsNeeded|foreach{
        write-progress -id 1 -activity "Creating MSOL tenant report" -status 'Running->' -percentcomplete ($i/$TotalFunctions*100) -currentOperation "Function: $_"
        & $_
        $i++
    }

If($OpenReport -eq "Yes"){
    [threading.thread]::CurrentThread.CurrentCulture = 'en-US'
	$Excel = New-Object -ComObject excel.application 
	$Excel.workbooks.open($ExcelReport) |out-null
	$Excel.Visible = $True

}

#---------------------------------------------------------[Script CleanUp]----------------------------------------------------------
#clean up all sessions
Remove-PSSession $SkypeSession
Remove-PSSession $ExOnSession

#remove all CSV reports
$csvReports.Values|Remove-Item

logwrite -logfile $Logfile -type Informational -Logstring "Script ended on $(hostname). The logfile is saved here: $($Logfile)"