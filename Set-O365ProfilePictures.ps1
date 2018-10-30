#-----------------------------------------------------------------------------
#
# Set-O365ProfilePicture
#
# This script update picture for each active Office 365 Exchange Online user
# should a picture be available in the specified directory
# 
#-----------------------------------------------------------------------------

#---------------------------
# Variables
#---------------------------

$PictureFolder = "C:\Images" #Without trailing slash
$PictureExt = "jpg"
$LogFolder = "c:\Temp\Logs\"

#---------------------------
# Log
#---------------------------
$LogTime = Get-Date -Format "dd.MM.yyyy hh-mm-ss"
#  Log file name:
$LogFile = $LogFolder + $LogTime + ".log"
#  Launch Transcript
Start-Transcript $LogFile

#---------------------------
# Main
#---------------------------

# Connection to Exchange Online with proxyMethod=RPS to upload images bigger that 10k
$UserCredential = Get-Credential 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS -Credential $UserCredential -Authentication Basic -AllowRedirection 
Import-PSSession $Session 

# Retreiving all active users
$UsersActive = Get-User -RecipientTypeDetails UserMailbox -ResultSize Unlimited | where {$_.UseraccountControl -notlike “*accountdisabled*”}

# Checking if there is a picture for each active user and uploading it if so
foreach($User in $UsersActive)
    {
        $PicturePath = $PictureFolder + "\" + $User.UserPrincipalName + "." + $PictureExt

        write-host "Looking for an image for : " + $User.DisplayName
       
        if(Test-Path $PicturePath)
            {
                write-host "Image exists : " + $PicturePath
        
                # INSERT IMAGE RESIZING HERE
            
                write-host "Updating " $User.DisplayName " image with: " $PicturePath
                Set-UserPhoto -Identity $User.UserPrincipalName -PictureData ([System.IO.File]::ReadAllBytes($PicturePath)) -Confirm:$false 
            }
    }

#  End of Log
Stop-Transcript 