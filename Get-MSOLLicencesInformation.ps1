$LogFile = "Office_365_Licenses.xls"
Import-Module MSOnline 
Connect-MsolService -Credential $Office365credentials 
 
write-host "Connecting to Office 365..." 
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1} 
foreach ($license in $licensetype)  
{  $headerstring = "DisplayName,UserPrincipalName,AccountSku"    
foreach ($row in $($license.ServiceStatus))  
    { 
        $headerstring = ($headerstring + "," + $row.ServicePlan.servicename) 
    } 
     
    Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append 
     
    write-host ("Gathering users with the following subscription: " + $license.accountskuid) 
	
	$users = Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses.accountskuid -contains $license.accountskuid} 
 foreach ($user in $users) { 
         
        write-host ("Processing " + $user.displayname) 
 
        $thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid} 
 
        $datastring = ($user.displayname + "," + $user.userprincipalname + "," + $license.SkuPartNumber) 
         
        foreach ($row in $($thislicense.servicestatus)) { 
             
            # Build data string 
            $datastring = ($datastring + "," + $($row.provisioningstatus)) 
        } 
         
        Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append 
    } 
 
    Out-File -FilePath $LogFile -InputObject " " -Encoding UTF8 -append 
}             
 
write-host ("Script Completed.  Results available in " + $LogFile)