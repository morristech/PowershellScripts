param (
    [string]$ExportPath = "C:\Temp"
)

 Get-MailboxExportRequest -status failed | Get-MailboxExportRequestStatistics -IncludeReport | Format-List > $ExportPath + "\report.txt"