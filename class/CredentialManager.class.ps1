Class CredentialManager {

    <##====================================================================================
        GLOBAL CONFIGURATION
    ##===================================================================================#>  
    [string] $Alias
    [string] $Username
    [securestring] $Password

    <##====================================================================================
        CONSTRUCTORS
    ##===================================================================================#>      

    CredentialManager(){
        this.Username = "Unknown"
    }

    <##====================================================================================
        FUNCTIONS
    ##===================================================================================#> 

    [void] Load ([string] $alias) {
        $this.Alias = $alias;  
    }   

    [void] Save ([string] $alias) {
        $this.Alias = $alias;  
    }  

    [string] Get () {
        return $this.Username
    }

    [void] Set ([string]$username, [SecureString] $password) {
        $this.Username = $username;
        $this.Password = $password;
    }

}
