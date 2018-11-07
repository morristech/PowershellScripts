enum WinSCPConnectionType {
    SCP
    SFTP
}

Class WinSCP {

    <##====================================================================================
        GLOBAL CONFIGURATION
    ##===================================================================================#>  
    [string] $Hostname

    <##====================================================================================
        CONSTRUCTORS
    ##===================================================================================#>      

    WinSCP(){
        this.$Hostname = "Unknown"
    }

    <##====================================================================================
        FUNCTIONS
    ##===================================================================================#> 

    [void] Connect ([string] $alias) {
        $this.Alias = $alias;  
    }  
    
    [void] GetAllFiles([bool] $overwrite) {
        
    }  

    [void] GetAllFilesOverwrite([string] $alias) {
        $this.GetAllFiles($true)
    }  

    [string] GetAllFilesSkip () {
        $this.GetAllFiles($false)
    }

    [void] Set ([string]$username, [SecureString] $password) {
        $this.Username = $username;
        $this.Password = $password;
    }

}