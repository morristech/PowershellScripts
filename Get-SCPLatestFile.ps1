param (
    $localPath = "c:\Temp\",
    $remotePath = "/home/pi/"
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
 
        # Get list of files in the directory
        $directoryInfo = $session.ListDirectory($remotePath)
 
        # Select the most recent file
        $latest =
            $directoryInfo.Files |
            Where-Object { -Not $_.IsDirectory } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
 
        # Any file at all?
        if ($latest -eq $Null)
        {
            Write-Host "No file found"
            exit 1
        }
 
        # Download the selected file
        $session.GetFiles($session.EscapeFileMask($remotePath + $latest.Name), $localPath).Check()
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