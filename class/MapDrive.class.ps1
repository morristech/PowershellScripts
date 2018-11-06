enum MapDriveType {
    NFS
    SMB
}

Class MapDrive {

    <##====================================================================================
        GLOBAL CONFIGURATION
    ##===================================================================================#>   

    [MapDriveType]$type
    [String]$letter
    [String]$target
    [String]$username
    [String]$password

    <##====================================================================================
	    CONSTRUCTORS
    ##===================================================================================#>

    MapDrive(){


    }

    <##====================================================================================
	    METHODS
    ##===================================================================================#>


    [bool] isElevated() {
        return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    # Remap drive if the script is run in elevated mode
    # Source : https://gist.github.com/anderssonjohan/8d3f958f29b4ae5c7802
    [void] fixDriveInElevatedMode(){

        if( this.isElevated ) {
            net use | Where-Object { $_ -match ":\s+\\\\"  -and !$_.StartsWith("Unavailable") } | ForEach-Object {
                $tokens = $_.split(":")
                $psdrivename = $tokens[0][$tokens[0].length-1]
                $path = $tokens[1].trim().split(" ")[0].trim()

                if( !(get-psdrive | ?{ $_.Name -eq $psdrivename } )) {
                    write-host ( "Restoring PSDrive for {0}: {1}" -f $psdrivename, $path )
                    new-psdrive $psdrivename FileSystem $path | out-null
                }
            }
        }

    }

    [bool] mapNFS() {
        return $true;

    }

    [bool] mapSMB(){
        return $true;

    }

    [bool] mapPSDrive() {
        return $true;
    }


}







