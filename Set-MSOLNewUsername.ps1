$CurrentUserName = Read-Host "Current username (Email Address):" 
$NewUserName = Read-Host "New username (Email Address):"  

set-msoluserprincipalname -newuserprincipalname $NewUserName -userprincipalname $CurrentUserName