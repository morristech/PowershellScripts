#Global Values
$localDriveLetter = "Z:"
$localDriveFolder = ""
$localDriveTraget = ""
$localDriveTragetUsername = ""
$localDriveTragetPassword = ""
$remoteHostName = ""
$remotePortNumber = ""
$remoteUserName = ""
$remoteSshPrivateKeyPath = ""
$remotePrivateKeyPassphrase = ''
$remoteSshHostKeyFingerprint = ""
$remotePath = ""
$mask = "*.*"
$logPath = ""
$WinSCPnetDllLocation = ""

$localPath = $localDriveLetter + $localDriveFolder

# Function Add-LogMessage to Write on Host and Log at the same time
function Add-LogMessage ($logFile, $message)
{
    $logLine = (Get-Date -Format o) + " " + $message

    # Stripping trailing characters
    $logLine = $logLine -replace "`t|`n|`r",""
    $logLine = $logLine -replace " ;|; ",";"

    Add-Content $logFile $logLine
    Write-Host $logLine
}

Add-LogMessage -logFile $logPath -message "Backup started"

try
{
    If (!(Test-Path $localDriveLetter)) {
        Add-LogMessage -logFile $logPath -message ("Mapping Network Drive: " + $localDriveLetter)
        $net = new-object -ComObject WScript.Network
        $net.MapNetworkDrive($localDriveLetter, $localDriveTraget, $false, $localDriveTragetUsername, $localDriveTragetPassword)
    }
    else {
        Add-LogMessage -logFile $logPath -message "Map Drive already exists"
    }
}
catch [Exception]
{
    Add-LogMessage -logFile $logPath -message ("Error Mapping Drive: " + $_.Exception.Message)
    exit 1
}

try
{
    # Load WinSCP .NET assembly
    Add-Type -Path $WinSCPnetDllLocation
 
    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Scp 
        HostName = $remoteHostName
        PortNumber = $remotePortNumber
        UserName = $remoteUserName
        SshPrivateKeyPath = $remoteSshPrivateKeyPath
        PrivateKeyPassphrase = $remotePrivateKeyPassphrase
        SshHostKeyFingerprint = $remoteSshHostKeyFingerprint
    }
 
    Add-LogMessage -logFile $logPath -message "Establising SCP Session"
    $session = New-Object WinSCP.Session
 
    try
    {
        # Connect
        $session.Open($sessionOptions)
 
        $files = $session.EnumerateRemoteFiles($remotePath, $mask, [WinSCP.EnumerationOptions]::None)

        foreach ($fileInfo in $files)
        {
            if(![System.IO.File]::Exists($localPath + $fileInfo.Name)){
                Add-LogMessage -logFile $logPath -message ("Downloading: " + $fileInfo.FullName)
                $session.GetFiles($session.EscapeFileMask($fileInfo.FullName), $localPath + "\*").Check() 
            }
            
        }

        if (Test-Path $localDriveLetter) {
            Add-LogMessage -logFile $logPath -message ("Removing Network Drive: " + $localDriveLetter)
            net use $localDriveLetter /delete
        }
    }
    finally
    {
        Add-LogMessage -logFile $logPath -message "Closing SCP Session"
        
        # Disconnect, clean up
        $session.Dispose()

        Add-LogMessage -logFile $logPath -message "Backup Completed"
    }

    exit 0
}
catch [Exception]
{
    Add-LogMessage -logFile $logPath -message ("Error: " + $_.Exception.Message)

    If (Test-Path $localDriveLetter) {
        Add-LogMessage -logFile $logPath -message ("Removing Network Drive: " + $localDriveLetter)
        net use $localDriveLetter /delete
    }

    exit 1
}