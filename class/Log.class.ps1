enum LogType {
    Message
    Success
    Warning
    Error
}

Class Log {

    <##====================================================================================
        GLOBAL CONFIGURATION
    ##===================================================================================#>  

    [string] $OutputFile
    [string] $TimeFormat = "%Y-%m-%d / %T" # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-6#notes
    [bool] $WriteTime = $true
    [bool] $WriteToHost = $true

    hidden [string] $LogLineFormat = ""

    <##====================================================================================
        CONSTRUCTORS
    ##===================================================================================#>      

    Log ( [string]$OutputFile ) {
        $this.OutputFile = $OutputFile
    }

    <##====================================================================================
        FUNCTIONS
    ##===================================================================================#>  

    # Stripping trailing characters
    [string] CleanString ($line){
        $logLine = $line
        $logLine = $logLine -replace "`t|`n|`r",""
        $logLine = $logLine -replace " ;|; ",";"

        return $logLine
    }

    [void] WriteMultipleLines([string] $messages){
        $lines = $messages.Split("`n`r")

        Foreach ($line in $lines){
            this.Write($line, [LogType].Message, $this.WriteToHost, $this.WriteTime)
		}

    }

    [void] Write ([string]$message, [LogType] $type, [bool]$writetohost, [bool]$writetime) {

        $logLine = $message

        if($writetime){
            $logLine = (Get-Date -UFormat $this.TimeFormat) + " " + $logLine
        }
              
        $logLine = $this.CleanString($logLine)
    
        Add-Content $this.OutputFile $logLine

        if($writetohost){
            Write-Host $logLine
        } 
    }    
	
	[void] Write ([string]$message, [bool]$writetohost) {
        this.Write($message, [LogType].Message, $writetohost, $this.WriteTime)           
    }

    [void] Write ($message) {
        this.Write($message, [LogType].Message, $this.WriteToHost, $this.WriteTime)
    }

}