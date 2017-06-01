param (
    $localPath = "c:\Temp\",
    $remotePath = "/home/user/",
    $mask = "*.*"
)
 
try
{
    # Load WinSCP .NET assembly
    Add-Type -Path "lib\WinSCP\WinSCPnet.dll"
 
    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        #Protocol = [WinSCP.Protocol]::Sftp
        Protocol = [WinSCP.Protocol]::Scp 
        HostName = ""
        UserName = ""
        Password = ""
        GiveUpSecurityAndAcceptAnySshHostKey = "true" 
        #SshHostKeyFingerprint = "*"
        #SshHostKeyFingerprint = "ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx"
    }
 
    $session = New-Object WinSCP.Session
 
    try
    {
        # Connect
        $session.Open($sessionOptions)
 
        $files = $session.EnumerateRemoteFiles($remotePath, $mask, [WinSCP.EnumerationOptions]::None)

        foreach ($fileInfo in $files)
        {
            if(![System.IO.File]::Exists($localPath + $fileInfo.Name)){
                Write-Host ("Downloading {0} ..." -f $fileInfo.FullName)
                $session.GetFiles($session.EscapeFileMask($fileInfo.FullName), $localPath + "\*").Check() 
            }
            
        }
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }
 
    exit 0
}
catch [Exception]
{
    Write-Host ("Error: {0}" -f $_.Exception.Message)
    exit 1
}