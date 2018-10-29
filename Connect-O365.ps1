<#PSScriptInfo
 
.VERSION 1.4
 
.GUID a3515355-c4b6-4ab8-8fa4-2150bbb88c96
 
.AUTHOR Jos Verlinde [MSFT]
 
.COMPANYNAME Microsoft
 
.COPYRIGHT
 
.TAGS O365 RMS 'Exchange Online' 'SharePoint Online' 'Skype for Business' 'PnP-Powershell' 'Office 365'
   
.LICENSEURI
 
.PROJECTURI
 
.ICONURI
 
.EXTERNALMODULEDEPENDENCIES MSOnline, Microsoft.Online.SharePoint.PowerShell, SkypeOnlineConnector, AADRM, OfficeDevPnP.PowerShell.V16.Commands
 
.REQUIREDSCRIPTS
 
.EXTERNALSCRIPTDEPENDENCIES
 
.RELEASENOTES
v1.4 Corrected bug wrt compliance search, remove prior created remote powershell sessions
V1.3 Add dependend module information
V1.2 add try-catch for SPO PNP Powershell, as that is less common
V1.1 initial publication to scriptcenter
#>

<#
.Synopsis
   Connect to Office 365 and get ready to admin anything.
.DESCRIPTION
   Connect to Office 365 and most related services and get ready to admin anything.
   The commandlet support saving the credentials in a save manner so that it can be used in unattended files
   requirement and dependency is that the relevant modules are installed on the systenm the cmdlet is used on.
    
   -O365
   -Azure Ad
   -SharePoint Online and PNP-PowerShell
   
       
.EXAMPLE
   connect-O365 -Account 'admin@contoso.com' -SharePoint
.EXAMPLE
   connect-O365 -Account 'admin@contoso.com' -SPO -EXO -Skype -Compliance -AADRM
.EXAMPLE
   connect-O365 -Account 'admin@contoso.com' -Persist:$false
.EXAMPLE
   connect-O365 -Account 'admin@contoso.com'
   #retrieve credentials for use in other cmdlets
   $Creds = Get-myCreds 'admin@contoso.com'
    
#>

[CmdletBinding()]
[Alias("COL")]
[OutputType([int])]
Param
(
    # Accoutn to authenticate with
    [Parameter(Mandatory=$false,Position=0)]
    [string]$Account,
        
    # Save the account credentials for later use
    [Parameter(Mandatory=$false)]
    [switch]$Persist = $false, 

    #Azure AD
    [Parameter(Mandatory=$false)]
    [Alias("AzureAD")] 
    [switch]$AAD = $true, 

    #Azure
# [Parameter(Mandatory=$false)]
# [switch]$Azure = $false,

    #Exchange Online
    [Parameter(Mandatory=$false)]
    [Alias("EXO")] 
    [switch]$Exchange = $false, 

    #Skype Online
    [Parameter(Mandatory=$false)]
    [Alias("CSO")] 
    [switch]$Skype = $false, 
    #
    [Parameter(Mandatory=$false)]
    [Alias("SPO")] 
    [switch]$SharePoint = $false, 
        
    #Compliance center
    [Parameter(Mandatory=$false)]
    [switch]$Compliance = $false,

    #Azure Rights Management
    [Parameter(Mandatory=$false)]
    [Alias("AZRMS")] 
    [Alias("RMS")]
    [switch]$AADRM = $false
     
)

function global:Store-myCreds ($username){
    $Credential = Get-Credential -Credential $username
    MkDir "$env:USERPROFILE\Creds" -ea 0 | Out-Null
    $Credential.Password | ConvertFrom-SecureString | Set-Content $env:USERPROFILE\Creds\$USERNAME.txt
    return $Credential 
 }

function global:Get-myCreds ($UserName , [switch]$Persist){
    $Store = "$env:USERPROFILE\creds\$USERNAME.txt"
    if (Test-Path $store ) {            
        $Password = Get-Content $store | ConvertTo-SecureString
        $Credential = New-Object System.Management.Automation.PsCredential($UserName,$Password)
        return $Credential
    } else {
        if ($persist -and -not [string]::IsNullOrEmpty($UserName)) {
            $admincredentials  = Store-myCreds $UserName
            return $admincredentials
        } else {
            return Get-Credential -Credential $username
        }
    }
 }
 
$admincredentials = Get-myCreds $account -Persist:$Persist
if ($admincredentials -eq $null){ throw "A valid Tenant Admin Account is required." } 

if($Close) {
    write-verbose "Closing open sessions for Exchange Online and Compliance Center"
    #Close Existing (remote Powershell Sessions)

    Get-PSSession -Name "Exchange Online" -ea SilentlyContinue | Remove-PSSession 
    Get-PSSession -Name "Compliance Center"  -ea SilentlyContinue | Remove-PSSession 
    Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession 
}


if ( $AAD) {
    write-verbose "Connecting to Azure AD"
    #Imports the installed Azure Active Directory module.
    Import-Module MSOnline -Verbose:$false 
    if (-not (Get-Module MSOnline ) ) { Throw "Module not installed"}
    #Establishes Online Services connection to Office 365 Management Layer.
    Connect-MsolService -Credential $admincredentials
}

IF ($Skype ){
    write-verbose "Connecting to Skype Online"
    #Imports the installed Skype for Business Online services module.
    Import-Module SkypeOnlineConnector -Verbose:$false  -Force 

    #Remove prior Session
    Get-PSSession -Name "Skype Online" -ea SilentlyContinue| Remove-PSSession 

    #Create a Skype for Business Powershell session using defined credential.
    $SkypeSession = New-CsOnlineSession -Credential $admincredentials -Verbose:$false
    $SkypeSession.Name="Skype Online"

    #Imports Skype for Business session commands into your local Windows PowerShell session.
    Import-PSSession $lyncSession -AllowClobber -Verbose:$false

}


If ($SharePoint) {
    write-verbose "Connecting to SharePoint Online"
    if (!$AAD) {
        Throw "AAD Connection required"
    } else {
        #get tenant name for AAD Connection
        $tname= (Get-MsolDomain | ?{ $_.IsInitial -eq $true}).Name.Split(".")[0]
    }

    #Imports SharePoint Online session commands into your local Windows PowerShell session.
    Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking -Verbose:$false
    #lookup the tenant name based on the intial domain for the tenant
    Connect-SPOService -url https://$tname-admin.sharepoint.com -Credential $admincredentials

    try { 
        write-verbose "Connecting to SharePoint Online PNP"
        import-Module OfficeDevPnP.PowerShell.V16.Commands -DisableNameChecking -Verbose:$false
        Connect-SPOnline -Credential $admincredentials -url "https://$tname.sharepoint.com"
    } catch {}
}


if ($Exchange ) {
    write-verbose "Connecting to Exchange Online"

    #Remove prior Session
    Get-PSSession -Name "Exchange Online" -ea SilentlyContinue| Remove-PSSession 

    #Creates an Exchange Online session using defined credential.
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $admincredentials -Authentication "Basic" -AllowRedirection
    $ExchangeSession.Name = "Exchange Online"
    #This imports the Office 365 session into your active Shell.
    Import-PSSession $ExchangeSession -AllowClobber -Verbose:$false

}

if ($Compliance) {
    write-verbose "Connecting to the Unified Compliance Center"
    #Remove prior Session
    Get-PSSession -Name "Compliance Center" -ea SilentlyContinue| Remove-PSSession 

    $PSCompliance = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $AdminCredentials -Authentication Basic -AllowRedirection
    $PSCompliance.Name = "Compliance Center"
    Import-PSSession $PSCompliance -AllowClobber -Verbose:$false 

}


If ($AADRM) {
    write-verbose "Connecting to Azure Rights Management"    
    #Azure RMS

    import-module AADRM -Verbose:$false
    Connect-AadrmService -Credential $admincredentials 

}
<#
if ($false) {
    #Azure MFA
    $auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $auth.RelyingParty = "*"
 
    #any devices issued for a user before this date would require MFA setup. Normally, we would select the date of running the command.
    $auth.RememberDevicesNotIssuedBefore = (Get-Date)
    $auth.State = "Enabled"
    Set-MsolUser -UserPrincipalName jos@atticware.com -StrongAuthenticationRequirements $auth -MobilePhone +31651446844
 
    $u = Get-MsolUser -UserPrincipalName jos@atticware.com
    $u | fl *
 
    $u.StrongAuthenticationPhoneAppDetails | FL *
 
    $u.StrongAuthenticationRequirements| FL *
    $u.StrongAuthenticationUserDetails| FL *
    $u.StrongAuthenticationProofupTime| FL *
 
    #MFA - ADFS server implements MFS - requires ADFS and on-prem MFA Server
    Set-MsolDomainFederationSettings -DomainName atticware.com -SupportsMFA $true
    get-MsolDomainFederationSettings -DomainName atticware.com
}
 
 
#>