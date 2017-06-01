# https://powershell.org/forums/topic/catch-too-long-file-paths/

$FileListDump = Get-ChildItem -Path C:\ -Recurse -ErrorAction SilentlyContinue -ErrorVariable err

foreach ($errorRecord in $err)
{
    if ($errorRecord.Exception -is [System.IO.PathTooLongException])
    {
        Write-Host $($errorRecord.TargetObject)
    }
}