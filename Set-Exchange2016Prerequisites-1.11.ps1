####################################################################################################################################################################
#  SCRIPT DETAILS                                                                                                                                                  #
#    Installs all required prerequisites for Exchange 206 for Windows Server 2012 (R2) components or Windows Server 2016,                                          #
#        downloading latest Update Rollup, etc.                                                                                                                    #
#																																								   #
# SCRIPT VERSION HISTORY																																		   #
#    Current Version	: 1.10																																	   #
#    Change Log			: 1.10 - Corrected .NET install process, fixed some registry entry work for RC4 disabling, added hotfix for Windows 2016                   #
#                       : 1.9 - Correct some additional coding errors, added PageFile Configuration																   #
#						: 1.8 - Completed recode of script and correcting bugs, typos and mode. Cleaned out old and duplicate code.								   #
#                       : 1.7 - Added Windows Server 2016 support (CU 3+)																						   #
#                       : 1.6 - Added hotfix for .NET 4.6.1 (required for Exchange)																				   #
#				        : 1.5 - Tweaked the script to allow .NET 4.5.2 or 4.6.1.  Added code in checker for .NET version and added individual installs for .NET	   #
#				        : 1.4 - Added .NET 4.6.1 installer for Exchange Server 2016 CU2 and higher																   #
#				        : 1.3 - Added SSL Security enhancements (optional)																						   #
#				        : 1.2 - Added High Performance Power Plan change, cleaned up menu																		   #
#				        : 1.1 - Added NIC Power Management																										   #
#				        : 1.0 - First iteration																													   #
#																																								   #
# OTHER SCRIPT INFORMATION																																		   #
#    Wish list			: Better comment based help																												   #
#				        : Event Log logging																														   #
#    Rights Required	: Local admin on server																												       #
#    Exchange Version	: 2016																																	   #
#    Author       		: Damian Scoles 																											  			   #
#    My Blog			: http://justaucguy.wordpress.com																										   #
#    Disclaimer   		: You are on your own.  This was not written by, supported by, or endorsed by Microsoft.												   #
#    Info Stolen from 	: Anderson Patricio, Bhargav Shukla and Pat Richard [Exchange 2010 script]																   #
#    					: http://msmvps.com/blogs/andersonpatricio/archive/2009/11/13/installing-exchange-server-2010-pre-requisites-on-windows-server-2008-r2.aspx#
#						: http://www.bhargavs.com/index.php/powershell/2009/11/script-to-install-exchange-2010-pre-requisites-for-windows-server-2008-r2/		   #
# 						: SQL Soldier - http://www.sqlsoldier.com/wp/sqlserver/enabling-high-performance-power-plan-via-powershell								   #
#																																								   #
# EXECUTION																																						   #
#. \Set-Exchange2016Prerequisites-1.10.ps1																														   #
#																																								   #
####################################################################################################################################################################

##################################
#   Global Variable Definitions  #
##################################

$Ver = (Get-WMIObject win32_OperatingSystem).Version
$OSCheck = $false
$Choice = "None"
$Date = get-date -Format "MM.dd.yyyy-hh.mm-tt"
$DownloadFolder = "c:\install"
$CurrentPath = (Get-Item -Path ".\" -Verbose).FullName
$Reboot = $false
$Error.clear()
Start-Transcript -path "$CurrenPath\$date-Set-Prerequisites.txt" | Out-Null
Clear-Host
# Pushd

############################################################
#   Global Functions - Shared between 2012 (R2) and 2016   #
############################################################

# Begin BITSCheck function
function BITSCheck {
    $Bits = Get-Module BitsTransfer
    if ($Bits -eq $null) {
        Write-Host "Importing the BITS module." -ForegroundColor cyan
        try {
            Import-Module BitsTransfer -erroraction STOP
        } catch {
            Write-Host "Server Management module could not be loaded." -ForegroundColor Red
        }
    }
} # End BITSCheck function

# Begin ModuleStatus function
function ModuleStatus {
        $module = Get-Module -name "ServerManager" -erroraction STOP

    if ($module -eq $null) {
        try {
            Import-Module -Name "ServerManager" -erroraction STOP
            # return $null
        } catch {
            Write-Host " ";Write-Host "Server Manager module could not be loaded." -ForegroundColor Red
        }
    } else {
        # Write-Host "Server Manager module is already imported." -ForegroundColor Cyan
        # return $null
    }
    Write-Host " "
} # End ModuleStatus function

# Begin FileDownload function
function FileDownload {
    param ($sourcefile)
    $Internetaccess = (Get-NetConnectionProfile -IPv4Connectivity Internet).ipv4connectivity
    If ($Internetaccess -eq "Internet") {
        if (Test-path $DownloadFolder) {
            Write-Host "Target folder $DownloadFolder exists." -ForegroundColor White
        } else {
            New-Item $DownloadFolder -type Directory | Out-Null
        }
        BITSCheck
        [string] $DownloadFile = $sourcefile.Substring($sourcefile.LastIndexOf("/") + 1)
        if (Test-Path "$DownloadFolder\$DownloadFile"){
            Write-Host "The file $DownloadFile already exists in the $DownloadFolder folder." -ForegroundColor Cyan
        } else {
            Start-BitsTransfer -Source "$SourceFile" -Destination "$DownloadFolder\$DownloadFile"
        }
    } else {
        Write-Host "This machine does not have internet access and thus cannot download required files. Please resolve!" -ForegroundColor Red
    }
} # End FileDownload function

# Configure the Server for the High Performance power plan
function highperformance {
    Write-Host " "
	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -ne $HighPerf) {
		powercfg -setactive $HighPerf
		CheckPowerPlan
	} else {
		if ($CurrPlan -eq $HighPerf) {
			Write-Host " ";Write-Host "The power plan is already set to " -nonewline;Write-Host "High Performance." -foregroundcolor green;Write-Host " "
		}
	}
}

# Check the server power management
function CheckPowerPlan {
	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -eq $HighPerf) {
		Write-Host " ";Write-Host "The power plan now is set to " -nonewline;Write-Host "High Performance." -foregroundcolor green;Write-Host " "
	}
}

# Turn off NIC power management
function PowerMgmt {
    Write-Host " "
	$NICs = Get-WmiObject -Class Win32_NetworkAdapter|Where-Object{$_.PNPDeviceID -notlike "ROOT\*" -and $_.Manufacturer -ne "Microsoft" -and $_.ConfigManagerErrorCode -eq 0 -and $_.ConfigManagerErrorCode -ne 22} 
	Foreach($NIC in $NICs) {
		$NICName = $NIC.Name
		$DeviceID = $NIC.DeviceID
		If([Int32]$DeviceID -lt 10) {
			$DeviceNumber = "000"+$DeviceID 
		} Else {
			$DeviceNumber = "00"+$DeviceID
		}
		$KeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\$DeviceNumber"
  
		If(Test-Path -Path $KeyPath) {
			$PnPCapabilities = (Get-ItemProperty -Path $KeyPath).PnPCapabilities
            # Check to see if the value is 24 and if not, set it to 24
            If($PnPCapabilities -ne 24){Set-ItemProperty -Path $KeyPath -Name "PnPCapabilities" -Value 24 | Out-Null}
            # Verify the value is now set to or was set to 24
			If($PnPCapabilities -eq 24) {Write-Host " ";Write-Host "Power Management has already been " -NoNewline;Write-Host "disabled" -ForegroundColor Green;Write-Host " "}
   		 } 
 	 } 
 }

 # Disable RC4
function DisableRC4 {
    Write-Host " "
	# Define Registry keys to look for
	$base = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\" -erroraction silentlycontinue
	$val1 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\" -erroraction silentlycontinue
	$val2 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\" -erroraction silentlycontinue
	$val3 = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\" -erroraction silentlycontinue
	
	# Define Values to add
	$registryBase = "Ciphers"
	$registryPath1 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128\"
	$registryPath2 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 40/128\"
	$registryPath3 = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 56/128\"
	$Name = "Enabled"
	$value = "0"
	$ssl = 0
	$checkval1 = Get-Itemproperty -Path "$registrypath1" -name $name -erroraction silentlycontinue
	$checkval2 = Get-Itemproperty -Path "$registrypath2" -name $name -erroraction silentlycontinue
	$checkval3 = Get-Itemproperty -Path "$registrypath3" -name $name -erroraction silentlycontinue
    
# Formatting for output
	Write-Host " "

# Add missing registry keys as needed
	If ($base -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL", $true)
		$key.CreateSubKey('Ciphers')
		$key.Close()
	} else {
		Write-Host "The " -nonewline;Write-Host "Ciphers" -ForegroundColor green -NoNewline;Write-Host " Registry key already exists."
	}

	If ($val1 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 128/128')
		$key.Close()
	} else {
		Write-Host "The " -nonewline;Write-Host "Ciphers\RC4 128/128" -ForegroundColor green -NoNewline;Write-Host " Registry key already exists."
	}

	If ($val2 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 40/128')
		$key.Close()
		New-ItemProperty -Path $registryPath2 -Name $name -Value $value -force -PropertyType DWord
	} else {
		Write-Host "The " -nonewline
        Write-Host "Ciphers\RC4 40/128" -ForegroundColor green -NoNewline
        Write-Host " Registry key already exists."
	}

	If ($val3 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 56/128')
		$key.Close()
	} else {
		Write-Host "The " -nonewline
        Write-Host "Ciphers\RC4 56/128" -ForegroundColor green -NoNewline
        Write-Host " Registry key already exists."
	}
	
# Add the enabled value to disable RC4 Encryption
	If ($checkval1.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath1 -Name $name -Value $value -force -PropertyType DWord
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		Write-Host "The registry value " -nonewline
        Write-Host "Enabled" -ForegroundColor green -NoNewline
        Write-Host " exists under the RC4 128/128 Registry Key."
        $ssl++
	}
	If ($checkval2.enabled -ne "0") {
		Write-Host $checkval2
		try {
			New-ItemProperty -Path $registryPath2 -Name $name -Value $value -force -PropertyType DWord
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		Write-Host "The registry value " -nonewline
        Write-Host "Enabled" -ForegroundColor green -NoNewline
        Write-Host " exists under the RC4 40/128 Registry Key."
        $ssl++
	}
	If ($checkval3.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath3 -Name $name -Value $value -force -PropertyType DWord
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		Write-Host "The registry value " -nonewline
        Write-Host "Enabled" -ForegroundColor green -NoNewline
        Write-Host " exists under the RC4 56/128 Registry Key."
        $ssl++
	}

# SSL Check totals
	If ($ssl -eq "3") {
		Write-Host " "
        Write-Host "RC4 " -ForegroundColor yellow -NoNewline
        Write-Host "is completely disabled on this server."
        Write-Host " "
	} 
	If ($ssl -lt "3"){
		Write-Host " "
        Write-Host "RC4 " -ForegroundColor yellow -NoNewline
        Write-Host "only has $ssl part(s) of 3 disabled.  Please check the registry to manually to add these values"
        Write-Host " "
	}
} # End of Disable RC4 function

# Disable SSL 3.0
function DisableSSL3 {
    Write-Host " "
    $TestPath1 = Get-Item -Path "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0" -erroraction silentlycontinue
    $TestPath2 = Get-Item -Path "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server" -erroraction silentlycontinue
    $registrypath = "HKLM:\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server"
    $Name = "Enabled"
	$value = "0"
    $checkval1 = Get-Itemproperty -Path "$registrypath" -name $name -erroraction silentlycontinue

# Check for SSL 3.0 Reg Key
	If ($TestPath1 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols", $true)
		$key.CreateSubKey('SSL 3.0')
		$key.Close()
	} else {
		Write-Host "The " -nonewline
        Write-Host "SSL 3.0" -ForegroundColor green -NoNewline
        Write-Host " Registry key already exists."
	}

# Check for SSL 3.0\Server Reg Key
	If ($TestPath2 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0", $true)
		$key.CreateSubKey('Server')
		$key.Close()
	} else {
		Write-Host "The " -nonewline
        Write-Host "SSL 3.0\Servers" -ForegroundColor green -NoNewline
        Write-Host " Registry key already exists."
	}

# Add the enabled value to disable SSL 3.0 Support
	If ($checkval1.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath -Name $name -Value $value -force -PropertyType DWord
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		Write-Host "The registry value " -nonewline
        Write-Host "Enabled" -ForegroundColor green -NoNewline
        Write-Host " exists under the SSL 3.0\Server Registry Key."
	}
} # End of Disable SSL 3.0 function

# Function - Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-WinUniComm4 {
    Write-Host " "
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		if($val.DisplayVersion -ne "5.0.8132.0"){
			if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is not installed.  Downloading and installing now." -foregroundcolor yellow
				Install-NewWinUniComm4
			} else {
    				Write-Host "`nAn old version of Microsoft Unified Communications Managed API 4.0 is installed."
				UnInstall-WinUniComm4
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now."  -foregroundcolor green
				Install-NewWinUniComm4
			}
   		} else {
   			Write-Host "`nThe Preview version of Microsoft Unified Communications Managed API 4.0 is installed."
   			UnInstall-WinUniComm4
   			Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now." -foregroundcolor green
   			Install-NewWinUniComm4
		}
	} else {
		Write-Host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
		Write-Host "installed." -ForegroundColor green
	}
    Write-Host " "
} # end Install-WinUniComm4

# Install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-NewWinUniComm4{
	FileDownload "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
	Set-Location $DownloadFolder
    [string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $targetfolder\WinUniComm4.log"
	Write-Host "File: UcmaRuntimeSetup.exe installing..." -NoNewLine
	Invoke-Expression $expression
	Start-Sleep -Seconds 20
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is now installed" -ForegroundColor Green
	}
    Write-Host " "
} # end Install-NewWinUniComm4

# Configure PageFile for Exchange
function ConfigurePageFile {
    $Stop = $False
    $WMIQuery = $False

    # Remove Existing PageFile
    try {
        Set-CimInstance -Query “Select * from win32_computersystem” -Property @{automaticmanagedpagefile=”False”}
    } catch {
        Write-Host "Cannot remove the existing pagefile." -ForegroundColor Red
        $WMIQuery = $True
    }
    # Remove PageFile with WMI if CIM fails
    If ($WMIQuery) {
		Try {
			$CurrentPageFile = Get-WmiObject -Class Win32_PageFileSetting
            $name = $CurrentPageFile.Name
            $CurrentPageFile.delete()
		} catch {
			Write-Host "The server $server cannot be reached via CIM or WMI." -ForegroundColor Red
			$Stop = $True
		}
    }

    # Get RAM and set ideal PageFileSize
    $GB = 1048576

    try {
        $RamInMb = (Get-CIMInstance -computername $name -Classname win32_physicalmemory -ErrorAction Stop | measure-object -property capacity -sum).sum/$GB
        $ExchangeRAM = $RAMinMb + 10
        # Set maximum pagefile size to 32 GB + 10 MB
        if ($ExchangeRAM -gt 32778) {$ExchangeRAM = 32778}
    } catch {
        Write-Host "Cannot acquire the amount of RAM in the server." -ForegroundColor Red
        $stop = $true
    }
    # Get RAM and set ideal PageFileSize - WMI Method
    If ($WMIQuery) {
		Try {
            $RamInMb = (Get-wmiobject -computername $server -Classname win32_physicalmemory -ErrorAction Stop | measure-object -property capacity -sum).sum/$GB
            $ExchangeRAM = $RAMinMb + 10

            # Set maximum pagefile size to 32 GB + 10 MB
            if ($ExchangeRAM -gt 32778) {$ExchangeRAM = 32778}
		} catch {
			Write-Host "Cannot acquire the amount of RAM in the server with CIM or WMI queries." -ForegroundColor Red
			$stop = $true
		}
    }

    # Reset WMIQuery
    $WMIQuery = $False

    if ($stop -ne $true) {
        # Configure PageFile
        try {
            Set-CimInstance -Query “Select * from win32_PageFileSetting” -Property @{InitialSize=$ExchangeRAM;MaximumSize=$ExchangeRAM}
        } catch {
            Write-Host "Cannot configure the PageFile correctly." -ForegroundColor Red
        }
        If ($WMIQuery) {
		    Try {
                Set-WMIInstance -computername $server -class win32_PageFileSetting -arguments @{name ="$name";InitialSize=$ExchangeRAM;MaximumSize=$ExchangeRAM}
		    } catch {
			    Write-Host "Cannot configure the PageFile correctly." -ForegroundColor Red
                $stop = $true
		    }
        }
        if ($stop -ne $true) {
            $pagefile = Get-CimInstance win32_PageFileSetting -Property * | select-object Name,initialsize,maximumsize
            $name = $pagefile.name;$max = $pagefile.maximumsize;$min = $pagefile.initialsize
            Write-Host " "
            Write-Host "This server's pagefile, located at " -ForegroundColor white -NoNewline
            Write-Host "$name" -ForegroundColor green -NoNewline
            Write-Host ", is now configured for an initial size of " -ForegroundColor white -NoNewline
            Write-Host "$min MB " -ForegroundColor green -NoNewline
            Write-Host "and a maximum size of " -ForegroundColor white -NoNewline
            Write-Host "$max MB." -ForegroundColor Green
            Write-Host " "
        } else {
            Write-Host "The PageFile cannot be configured at this time." -ForegroundColor Red
        }
    } else {
        Write-Host "The PageFile cannot be configured at this time." -ForegroundColor Red
    }
}

######################################################
#    This section is for the Windows 2012 (R2) OS    #
######################################################

function Code2012 {

# Start code block for Windows 2012 or 2012 R2

$Menu2012 = {

    Write-Host "	********************************************************************" -ForegroundColor Cyan
    Write-Host "	 Exchange Server 2016 [On Windows 2012 (R2)] - Prerequisites script" -ForegroundColor Cyan
    Write-Host "	********************************************************************" -ForegroundColor Cyan
    Write-Host " "
    Write-Host "	.NET UPDATE - Added ALL hotfixes for .NET 4.6.1+ and added .NET 4.6.2 install option" -ForegroundColor Red
    Write-Host " "
    Write-Host "	Please select an option from the list below:" -ForegroundColor White
    Write-Host "	"
    Write-Host "	CU3+ [.NET 4.6.2]" -ForegroundColor Yellow -NoNewLine
    Write-Host " - RECOMMENDED" -ForegroundColor Green
    Write-Host "	1) Install Mailbox prerequisites - Part 1" -ForegroundColor White
    Write-Host "	2) Install Mailbox prerequisites - Part 2" -ForegroundColor White
    Write-Host "	3) Install Edge Transport prerequisites" -ForegroundColor White
    Write-Host "	"
    Write-Host "	CU2 to CU4 [.NET 4.6.1]" -ForegroundColor Yellow -NoNewLine
    Write-Host " - Will be removed in a future version." -ForegroundColor Cyan
    Write-Host "	4) Install Mailbox prerequisites - Part 1" -ForegroundColor White
    Write-Host "	5) Install Mailbox prerequisites - Part 2" -ForegroundColor White
    Write-Host "	6) Install Edge Transport prerequisites" -ForegroundColor White
    Write-Host " "
    Write-Host "    ** No option for .NET 4.5.2 - REMOVED **" -ForegroundColor Red
    Write-Host " "
    Write-Host "	10) Launch Windows Update" -ForegroundColor White
    Write-Host "	11) Check Prerequisites for Mailbox role" -ForegroundColor White
    Write-Host "	12) Check Prerequisites for Edge role" -ForegroundColor White
    Write-Host "	"
    Write-Host "	20) Install - One-Off - .NET 4.5.2 - RTM or CU1" -ForegroundColor White
    Write-Host "	21) Install - One-Off - Prereq For .Net 4.6.1 - CU2+" -ForegroundColor White
    Write-Host "	22) Install - One-Off - .NET 4.6.1 - CU2+" -ForegroundColor White
    Write-Host "	23) Install - One-Off - .NET 4.6.2 - CU3+" -ForegroundColor White
    Write-Host "	24) Install - One-Off - Windows Features [MBX]" -ForegroundColor White
    Write-Host "	25) Install - One Off - Unified Communications Managed API 4.0" -ForegroundColor White
    Write-Host "	"
    Write-Host "	30) Set Power Plan to High Performance" -ForegroundColor White
    Write-Host "	31) Disable Power Management for NICs." -ForegroundColor White
    Write-Host "	32) Disable SSL 3.0 Support" -ForegroundColor White
    Write-Host "	33) Disable RC4 Support" -ForegroundColor White
    Write-Host "	"
    Write-Host "	POST EXCHANGE 2016 INSTALL" -ForegroundColor Yellow
    Write-Host "	40) Configure PageFile to RAM + 10 MB" -foregroundcolor green
    Write-Host "	"
    Write-Host "	98) Restart the Server"  -ForegroundColor Red
    Write-Host "	99) Exit" -ForegroundColor Cyan
    Write-Host "	"
    Write-Host "	Select an option.. [1-99]? " -ForegroundColor White -nonewline
}

# Function - Pre .NET 4.6.1
Function Install-PreNET461 {
    # Verify .NET 4.6.1 is not already installed
    Check-DotNetVersion
    $DotNetVersion = ($global:NetVersion).release
    if ($DotNetVersion -lt 394271) {
        Write-Host " "
        Write-Host ".NET 4.6.1 is not installed." -ForegroundColor Yellow
        Write-Host "Installing all patches required in order to install .NET 4.6.1." -ForegroundColor White
        Write-Host " "

        # KB2919442 Install
        FileDownload "https://download.microsoft.com/download/D/6/0/D60ED3E0-93A5-4505-8F6A-8D0A5DA16C8A/Windows8.1-KB2919442-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        Write-Host "File: Windows8.1-KB2919442-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2919442-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2919442 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true

        # clearcompressionflag.exe Install
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/clearcompressionflag.exe"
        Set-Location $DownloadFolder
        # [string]$expression = ".\clearcompressionflag.exe /quiet /norestart /l* $DownloadFolder\clearcompressionflag.log"
        [string]$expression = ".\clearcompressionflag.exe /norestart /l* $DownloadFolder\clearcompressionflag.log"
        write-host " "
        Write-Host "File: clearcompressionflag.exe installing..." -NoNewLine
        Invoke-Expression $expression
        Start-Sleep -Seconds 60
        Write-Host "`nClearcompressionflag.exe has been run" -ForegroundColor Green
        write-host " "
        $Reboot = $true

        # KB2919355 Install
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2919355-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        write-host "WARNING: THIS HOTFIX COULD TAKE 25 MINUTES TO INSTALL!!!" -ForegroundColor Red
        write-host " "
        Write-Host "File: Windows8.1-KB2919355-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2919355-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2919355 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true

        # KB2932046 Install
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2932046-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        Write-Host "File: Windows8.1-KB2932046-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2932046-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2932046 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true

        # KB2959977 Install    
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2959977-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        Write-Host "File: Windows8.1-KB2959977-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2959977-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2959977 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true    

        # KB2937592 Install    
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2937592-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        Write-Host "File: Windows8.1-KB2937592-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2937592-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2937592 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true    

        # KB2938439 Install
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2938439-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        Write-Host "File: Windows8.1-KB2938439-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2938439-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2938439 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true

        # KB2934018 Install    
        FileDownload "https://download.microsoft.com/download/2/5/6/256CCCFB-5341-4A8D-A277-8A81B21A1E35/Windows8.1-KB2934018-x64.msu"
        Set-Location $DownloadFolder
        write-host " "
        Write-Host "File: Windows8.1-KB2934018-x64.msu installing..." -NoNewLine
        $HotFixInstall={
            $arglist='Windows8.1-KB2934018-x64.msu','/quiet','/norestart'
            Start-Process -FilePath 'c:\windows\system32\wusa.exe' -ArgumentList $arglist -NoNewWindow -Wait
        }
        Invoke-Command -ScriptBlock $HotFixInstall
        Start-Sleep -Seconds 60
        Write-Host "`nKB2934018 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true
    }
    start-sleep 2
} # End PreNET461 function

# Function - .NET 4.6.1
function Install-DotNET461 {
    # Verify .NET 4.6.1 is not already installed
    Check-DotNetVersion
    $DotNetVersion = ($global:NetVersion).release
    if ($DotNetVersion -lt 394271) {
        # Download .NET 4.6.1 installer
        FileDownload "https://download.microsoft.com/download/E/4/1/E4173890-A24A-4936-9FC9-AF930FE3FA40/NDP461-KB3102436-x86-x64-AllOS-ENU.exe"
	    Set-Location $DownloadFolder
        [string]$expression = ".\NDP461-KB3102436-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $DownloadFolder\DotNET461.log"
        Write-Host " "
	    Write-Host "File: NDP461-KB3102436-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
        Invoke-Expression $expression
        Start-Sleep -Seconds 60
        Write-Host "`n.NET 4.6.1 is now installed" -ForegroundColor Green
        Write-Host " "
        $Reboot = $true
    } 
    start-sleep 2
}

# Function - Post .NET 4.6.1 - 2012
Function 2012PostNET461 {
    Check-DotNetVersion
    $DotNetVersion = ($global:NetVersion).release
    if ($DotNetVersion -eq 394271) {
	    # Download the Hotfix
        FileDownload "http://download.microsoft.com/download/E/F/1/EF1FB34B-58CB-4568-85EC-FA359387E328/Windows8-RT-KB3146714-x64.msu"
	    Set-Location $DownloadFolder
	    [string]$expression = "wusa.exe .\Windows8-RT-KB3146714-x64.msu /quiet /norestart"
        Write-Host " "
        Write-Host "File: Windows8-RT-KB3146714-x64.msu installing..." -NoNewLine
        Invoke-Expression $expression
        Start-Sleep -Seconds 60
        Write-Host "`n.HotFix KB3146714 is now installed" -ForegroundColor Green
        Write-Host " "
    } 
    start-sleep 2
} # Post .NET 4.6.1 - 2012 Function

# Function - Post .NET 4.6.1 - 2012 R2
Function 2012R2PostNET461 {
    Check-DotNetVersion
    $DotNetVersion = ($global:NetVersion).release
    if ($DotNetVersion -eq 394271) {
        # Download the hotfix
	    FileDownload "http://download.microsoft.com/download/E/F/1/EF1FB34B-58CB-4568-85EC-FA359387E328/Windows8.1-KB3146715-x64.msu"
	    Set-Location $DownloadFolder
	    [string]$expression = "wusa.exe .\Windows8.1-KB3146715-x64.msu /quiet /norestart"
        Write-Host " "
        Write-Host "File: Windows8.1-KB3146715-x64.msu installing..." -NoNewLine
        Invoke-Expression $expression
        Start-Sleep -Seconds 60
        Write-Host "`n.HotFix KB3146715 is now installed" -ForegroundColor Green
        Write-Host " "
    } 
    start-sleep 2
} # End Post .NET 4.6.1 - 2012 R2 Function

# Function - Post .NET 4.6.1
function POSTDotNET461 {
    if ((Get-WMIObject win32_OperatingSystem).Version -match '6.2'){
        2012PostNET461
    }
    
    if ((Get-WMIObject win32_OperatingSystem).Version -match '6.3'){
        2012R2PostNET461
    }
} # End Post .NET 4.6.1 Function

# Function - .NET 4.6.2 [for CU14+]
function Install-DotNET462 {
    Check-DotNetVersion
    $DotNetVersion = ($global:NetVersion).release
    if ($DotNetVersion -lt 394806) {
	    # Download .NET 4.6.2
        FileDownload "https://download.microsoft.com/download/F/9/4/F942F07D-F26F-4F30-B4E3-EBD54FABA377/NDP462-KB3151800-x86-x64-AllOS-ENU.exe"
	    Set-Location $DownloadFolder
        [string]$expression = ".\NDP462-KB3151800-x86-x64-AllOS-ENU.exe /quiet /norestart /l* $DownloadFolder\DotNET461.log"
        write-host " "
	    Write-Host "File: NDP462-KB3151800-x86-x64-AllOS-ENU.exe installing..." -NoNewLine
        Invoke-Expression $expression
        Start-Sleep -Seconds 60
        Write-Host "`n.NET 4.6.2 is now installed" -ForegroundColor Green
        write-host " "
        $Reboot = $true
     } 
    start-sleep 2
} # End Install-DotNET462

# Function - Check Dot Net Version
function Check-DotNetVersion {
    # Formatting
    Write-Host " "
    Write-Host " "
    # .NET 4.5.2 or 4.6.1
	$NETval = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name "Release"
	if ($NETval.Release -lt "379893") {
		write-host ".NET 4.5.2 is " -nonewline 
		write-host "not installed!" -ForegroundColor red -nonewline
        write-host " - this does not meet the minimum requirements for a support Exchange version to be installed." -ForegroundColor white
        write-host " "
	}

    if ($NETval.Release -eq "379893") {
    	write-host ".NET 4.5.2 is " -nonewline -ForegroundColor white
		write-host "installed." -ForegroundColor green -NoNewline
        write-host " - this is sufficient for any version of Exchange Server 2013 CU7+." -ForegroundColor white
        write-host " "
    }

    if ($NETval.Release -eq "394271") {
    	write-host ".NET 4.6.1 is " -nonewline -ForegroundColor white
		write-host "installed," -ForegroundColor green -nonewline
        write-host " but it is missing KB3146711." -ForegroundColor yellow  -nonewline
        write-host " Make sure to install that before installing Exchange 2013." -ForegroundColor white
        write-host " "
    }

    if ($NETval.Release -eq "394294") {
    	Write-Host ".NET 4.6.1 with KB3146711 is " -nonewline -ForegroundColor White
		Write-Host "installed." -ForegroundColor green -nonewline
        Write-Host " - This version of .NET is suitable for " -NoNewline -ForegroundColor White
        Write-Host "Exchange Server 2013 CU13+." -ForegroundColor Yellow
        Write-Host " "
    }

    if ($NETval.Release -eq "394806") {
    	Write-Host ".NET 4.6.2 is " -nonewline -ForegroundColor White
		Write-Host "installed." -ForegroundColor green -nonewline
        Write-Host " - This version of .NET is suitable for " -NoNewline -ForegroundColor White
        Write-Host "Exchange Server 2013 CU14+" -foregroundcolor yellow
        Write-Host " "
    }
    $global:NetVersion = $NETVal
} # end Check-DotNetVersion

# Mailbox Role - Windows Feature requirements
function check-MBXprereq {
    Write-Host " "
    Write-Host "Checking all requirements for the Mailbox Role in Exchange Server 2016....." -foregroundcolor yellow
    Write-Host " "
    start-sleep 2

# .NET Check
    Check-DotNetVersion

# Windows Feature Check
	$values = @("AS-HTTP-Activation","Desktop-Experience","NET-Framework-45-Features","RPC-over-HTTP-proxy","RSAT-Clustering","RSAT-Clustering-CmdInterface","RSAT-Clustering-Mgmt","RSAT-Clustering-PowerShell","Web-Mgmt-Console","WAS-Process-Model","Web-Asp-Net45","Web-Basic-Auth","Web-Client-Auth","Web-Digest-Auth","Web-Dir-Browsing","Web-Dyn-Compression","Web-Http-Errors","Web-Http-Logging","Web-Http-Redirect","Web-Http-Tracing","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Lgcy-Mgmt-Console","Web-Metabase","Web-Mgmt-Console","Web-Mgmt-Service","Web-Net-Ext45","Web-Request-Monitor","Web-Server","Web-Stat-Compression","Web-Static-Content","Web-Windows-Auth","Web-WMI","Windows-Identity-Foundation")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "installed." -ForegroundColor green
		}else{
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "not installed!" -ForegroundColor red
		}
	}

# Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit 
  $val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
  if($val.DisplayVersion -ne "5.0.8308.0"){
    	if($val.DisplayVersion -ne "5.0.8132.0"){
        	if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
			Write-Host "No version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
            		Write-Host "not installed!" -ForegroundColor red
            		Write-Host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
		} else {
			Write-Host "The Preview version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
			Write-Host "installed." -ForegroundColor red
			Write-Host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red
			Write-Host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
		}
	} else {
        	Write-Host "The wrong version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        	Write-Host "installed." -ForegroundColor red
        	Write-Host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red 
        	Write-Host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
        }   
   } else {
        Write-Host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        Write-Host "installed." -ForegroundColor green
   }
   Write-Host " "
   Write-Host " "
} # End function check-MBXprereq

# Edge Transport requirement check
function check-EdgePrereq {

	Write-Host " "
    Write-Host "Checking all requirements for the Edge Transport Role in Exchange Server 2016....." -foregroundcolor yellow
    Write-Host " "
    start-sleep 2

    # Check .NET version
    Check-DotNetVersion

     # Windows Feature AD LightWeight Services
	$values = @("ADLDS")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "installed." -ForegroundColor green
            Write-Host " "
		}else{
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "not installed!" -ForegroundColor red
            Write-Host " "
		}
	}
    Write-Host " "
} # End Check-EdgePrereq

# Install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-NewWinUniComm4{
	FileDownload "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
	Set-Location $DownloadFolder
	[string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $DownloadFolder\WinUniComm4.log"
	Write-Host "File: UcmaRuntimeSetup.exe installing..." -NoNewLine
	Invoke-Expression $expression
	Start-Sleep -Seconds 20
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is now installed" -ForegroundColor Green
	}
    Write-Host " "
} # end Install-NewWinUniComm4

Do { 	
	if ($Reboot -eq $true){Write-Host "`t`t`t`t`t`t`t`t`t`n`t`t`t`tREBOOT REQUIRED!`t`t`t`n`t`t`t`t`t`t`t`t`t`n`t`tDO NOT INSTALL EXCHANGE BEFORE REBOOTING!`t`t`n`t`t`t`t`t`t`t`t`t" -backgroundcolor red -foregroundcolor black}
	if ($Choice -ne "None") {Write-Host "Last command: "$Choice -foregroundcolor Yellow}	
    invoke-command -scriptblock $Menu2012
	$Choice = Read-Host

  switch ($Choice)    {
  ##### NEW OPTION LIST #####


##### .NET 4.6.2 [CU4+] Section #####
    1 { # Prep Mailbox Role - Part 1
        ModuleStatus -name ServerManager
        Install-DotNET462
        Install-WindowsFeature RSAT-ADDS
        Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
        highperformance
        PowerMgmt
        $Reboot = $true
    }
    2 { # Prep Mailbox Role - Part 2
        ModuleStatus -Name ServerManager
        Install-WinUniComm4
        $Reboot = $true
    }
    3 {# Prep Exchange Transport
        Install-windowsfeature ADLDS 
        Install-DotNET462
    }

##### .NET 4.6.1 [CU2+] Section  #####
    4 { # Prep Mailbox Role - Part 1
        ModuleStatus -Name ServerManager
        Install-PreNET461
        Install-WindowsFeature RSAT-ADDS
        Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
        Highperformance
        PowerMgmt
        $Reboot = $true
    }
    5 { # Prep Mailbox Role - Part 2
        Install-DotNET461
        POSTDotNET461
        ModuleStatus -Name ServerManager
        Install-WinUniComm4
        $Reboot = $true
    }
    6 {# Prep Edge Transport Role
        Install-windowsfeature ADLDS 
        Install-DotNET461
        POSTDotNET461
        $Reboot = $true
    }

##### All other options #####

    10 { #	Windows Update
        Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
    }
    11 { #	Mailbox Requirement Check
        Check-MBXprereq
    }
    12 { #	Edge Transport Requirement Check
        Check-EdgePrereq
    }
    20 { #	Install -One-Off - .NET 4.5.2 - RTM to CU4
        ModuleStatus -name ServerManager
        Install-DotNET452
    }
    21 { # Install - prereq for .Net 4.6.1
        Install-PreNET461
        $Reboot = $true
    }
    22 { # Install - .NET 4.6.1 - CU2 to CU4
        ModuleStatus -name ServerManager
        Install-DotNET461
        POSTDotNET461
        $Reboot = $true
    }
    23 { # Install - .NET 4.6.2 - CU3+
        ModuleStatus -name ServerManager
        Install-DotNET462
        $Reboot = $true
    }
    24 {#	Install -One-Off - Windows Features [MBX]
        ModuleStatus -name ServerManager
        Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
    }
    25 {#	Install - One Off - Unified Communications Managed API 4.0
        Install-WinUniComm4
    }
    30 { # Set power plan to High Performance as per Microsoft
        HighPerformance
    }
    31 { # Disable Power Management for NICs.		
        PowerMgmt
    }
    32 { # Disable SSL 3.0 Support
        DisableSSL3
    }
    33 { # Disable RC4 Support		
        DisableRC4
    }
    40 {#   Add Windows Defender Exclsions
        ConfigurePagefile
    }
    98 {#	Exit and restart
        Stop-Transcript
        Restart-Computer -ComputerName LocalHost -Force
    }
    99 {#	Exit
        if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
            Write-Host "BitsTransfer: Removing..." -NoNewLine
            Remove-Module BitsTransfer
            Write-Host "`b`b`b`b`b`b`b`b`b`b`bremoved!   " -ForegroundColor Green
        }
        popd
        Write-Host "Exiting..."
        Stop-Transcript
    }
    default {Write-Host "You haven't selected any of the available options. "}
  }
} while ($Choice -ne 99)
}

######################################################
#    This section is for the Windows 2016 OS         #
######################################################

function Code2016 {

    # Start code block for Windows 2016 Server

$Menu2016 = {

    Write-Host "	******************************************************************" -ForegroundColor Cyan
    Write-Host "	Exchange Server 2016 on Windows Server 2016 - Features script" -ForegroundColor Cyan
    Write-Host "	******************************************************************" -ForegroundColor Cyan
    Write-Host "		"
    Write-Host "	Please select an option from the list below:" -foregroundcolor green
    Write-Host "	"
    Write-Host "	1) Install Mailbox prerequisites - Part 1 - CU3+" -ForegroundColor White
    Write-Host "	2) Install Mailbox prerequisites - Part 2 - CU3+" -ForegroundColor White
    Write-Host "	3) Install Edge Transport Server prerequisites - CU3 +" -ForegroundColor White
    Write-Host "	"
    Write-Host "	10) Launch Windows Update" -ForegroundColor White
    Write-Host "	11) Check Prerequisites for Mailbox role" -ForegroundColor White
    Write-Host "	12) Check Prerequisites for Edge role" -ForegroundColor White
    Write-Host "	"
    Write-Host "	20) Install - One-Off - Windows Features [MBX]" -ForegroundColor White
    Write-Host "	21) Install - One Off - Unified Communications Managed API 4.0" -ForegroundColor White
    Write-Host "	"
    Write-Host "	30) Set Power Plan to High Performance" -ForegroundColor White
    Write-Host "	31) Disable Power Management for NICs." -ForegroundColor White
    Write-Host "	32) Disable SSL 3.0 Support" -ForegroundColor White
    Write-Host "	33) Disable RC4 Support" -ForegroundColor White
    Write-Host "	"
    Write-Host "	POST EXCHANGE 2016 INSTALL" -ForegroundColor Yellow
    Write-Host "	40) Add Windows Defender exclusions"
    Write-Host "	41) Configure PageFile to RAM + 10 MB" -foregroundcolor green
    Write-Host "	"
    Write-Host "	98) Restart the Server"  -ForegroundColor Red
    Write-Host "	99) Exit" -ForegroundColor Cyan
    Write-Host "	"
    Write-Host "	Select an option.. [1-99]? " -nonewline
}

# Check Mailox Requirements
function check-MBXprereq {
    Write-Host " "
    Write-Host "Checking all requirements for the Mailbox Role in Exchange Server 2016 on Windows Server 2016....." -foregroundcolor yellow
    Write-Host " "
    start-sleep 2

# .NET Check - Removed as Windows 2016 has .NET 4.6.2 by default

# Windows Feature Check
	$values = @("NET-Framework-45-Features","RPC-over-HTTP-proxy","RSAT-Clustering","RSAT-Clustering-CmdInterface","RSAT-Clustering-Mgmt","RSAT-Clustering-PowerShell","Web-Mgmt-Console","WAS-Process-Model","Web-Asp-Net45","Web-Basic-Auth","Web-Client-Auth","Web-Digest-Auth","Web-Dir-Browsing","Web-Dyn-Compression","Web-Http-Errors","Web-Http-Logging","Web-Http-Redirect","Web-Http-Tracing","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Lgcy-Mgmt-Console","Web-Metabase","Web-Mgmt-Console","Web-Mgmt-Service","Web-Net-Ext45","Web-Request-Monitor","Web-Server","Web-Stat-Compression","Web-Static-Content","Web-Windows-Auth","Web-WMI","Windows-Identity-Foundation")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "installed." -ForegroundColor green
		}else{
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "not installed!" -ForegroundColor red
		}
	}

# Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit 
  $val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
  if($val.DisplayVersion -ne "5.0.8308.0"){
    	if($val.DisplayVersion -ne "5.0.8132.0"){
        	if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
			Write-Host "No version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
            		Write-Host "not installed!" -ForegroundColor red
            		Write-Host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
		} else {
			Write-Host "The Preview version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
			Write-Host "installed." -ForegroundColor red
			Write-Host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red
			Write-Host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
		}
	} else {
        	Write-Host "The wrong version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        	Write-Host "installed." -ForegroundColor red
        	Write-Host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red 
        	Write-Host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
        }   
   } else {
        Write-Host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        Write-Host "installed." -ForegroundColor green
   }
} # End function check-MBXprereq

# Check Edge Requirements
function Check-EdgePrereq {

    Write-Host " "
    Write-Host "Checking all requirements for the Edge Transport Role in Exchange Server 2016 on Windows Server 2016....." -foregroundcolor yellow
    Write-Host " "
    start-sleep 2

    # Check .NET version - Removed as Windows 2016 has .NET 4.6.2 by default

    # Windows Feature AD LightWeight Services
	$values = @("ADLDS")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "installed." -ForegroundColor green
            Write-Host " "
		}else{
			Write-Host "The Windows Feature"$item" is " -nonewline 
			Write-Host "not installed!" -ForegroundColor red
            Write-Host " "
		}
	}
}

# Start Windows Defender function
function WindowsDefender {
    Write-Host " "
    Write-Host "Windows Defender exclusions:" -ForegroundColor cyan
    if (Get-Module Defender -ListAvailable) {
        try {
            # Noderunner exclusion
            $ExchangeProcess = "$exinstall\Bin\Search\Ceres\Runtime\1.0\Noderunner.exe"
            Add-MpPreference -ExclusionProcess $ExchangeProcess
            Write-Host "Added " -ForegroundColor White -nonewline
            Write-Host "Process exclusions" -foregroundcolor green -nonewline
            Write-Host " successfully!" -ForegroundColor White
        } catch {
            Write-Warning $_.Exception.Message
        }
        try {
            # System Drive
            $Drive = $env:SystemDrive
            $ExchangeSetupLog = "$drive\ExchangeSetupLogs\ExchangeSetup.log"
            Add-MpPreference -ExclusionPath $ExchangeSetupLog
            Write-Host "Added " -ForegroundColor White -nonewline
            Write-Host "Setup Log exclusions" -foregroundcolor green -nonewline
            Write-Host " successfully!" -ForegroundColor White
        } catch {
            Write-Warning $_.Exception.Message
        }
        try {
            # Exchange Installation Director
            Add-MpPreference -ExclusionPath $exinstall
            Write-Host "Added " -ForegroundColor White -nonewline
            Write-Host "Exchange Install directory exclusions" -foregroundcolor green -nonewline
            Write-Host " successfully!" -ForegroundColor White
        } catch {
            Write-Warning $_.Exception.Message
        }
    } else {
        Write-Warning "Windows Defender PowerShell module not available."
    }
    Write-Host " "
} # End Windows Defender function

Do { 	
	if ($Reboot -eq $true){Write-Host "`t`t`t`t`t`t`t`t`t`n`t`t`t`tREBOOT REQUIRED!`t`t`t`n`t`t`t`t`t`t`t`t`t`n`t`tDO NOT INSTALL EXCHANGE BEFORE REBOOTING!`t`t`n`t`t`t`t`t`t`t`t`t" -backgroundcolor red -foregroundcolor black}
	if ($Choice -ne "None") {Write-Host "Last command: "$Choice -foregroundcolor Yellow}	
	invoke-command -scriptblock $Menu2016
    $Choice = Read-Host

  switch ($Choice)    {
    1 {#	Prep Mailbox Role - Part 1 - CU3+
      ModuleStatus -name ServerManager
      Install-WindowsFeature RSAT-ADDS
      Install-WindowsFeature NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
      HighPerformance
      PowerMgmt
      $Reboot = $true
    }
    2 {#	Prep Mailbox Role - Part 2 - CU3+
      ModuleStatus -name ServerManager
      Install-WinUniComm4
      $Reboot = $true
    }
    3 {#	Prep Exchange Transport - CU3+
      Install-windowsfeature ADLDS
    }
    10 {#	Windows Update
      Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
    }
    11 {#	Mailbox Requirement Check
      Check-MBXprereq
    }
    12 {#	Edge Transport Requirement Check
      Check-EdgePrereq
    }
    20 {#	Install -One-Off - Windows Features [MBX] - CU3+
      ModuleStatus -name ServerManager
      Install-WindowsFeature NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
    }
    21 {#	Install - One Off - Unified Communications Managed API 4.0 - CU3+
      Install-WinUniComm4
    }
    30 {#	Set power plan to High Performance as per Microsoft
      HighPerformance
    }
    31 {#	Disable Power Management for NICs.		
      PowerMgmt
    }
    32 {#	Disable SSL 3.0 Support
      DisableSSL3
    }
    33 {#	Disable RC4 Support		
      DisableRC4
    }
    40 {#   Add Windows Defender Exclsions
        WindowsDefender
    }
    41 {#   Configure the pagefile to be RAM + 10 and not system managed
        ConfigurePageFile
    }
    98 {#	Exit and restart
      Stop-Transcript
      restart-computer -computername localhost -force
    }
    99 {#	Exit
      if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
        Write-Host "BitsTransfer: Removing..." -NoNewLine
        Remove-Module BitsTransfer
        Write-Host "`b`b`b`b`b`b`b`b`b`b`bremoved!   " -ForegroundColor Green
      }
      popd
      Write-Host "Exiting..."
      Stop-Transcript
    }
    default {Write-Host "You haven't selected any of the available options. "}
  }
} while ($Choice -ne 99)

} 

######################################################
#               MAIN SCRIPT BODY                     #
######################################################

# Check for Windows 2012 or 2012 R2
if (($ver -match '6.2') -or ($ver -match '6.3')) {
    $OSCheck = $true
    Code2012
}

# Check for Windows 2016
if ($ver -match '10.0') {
    $OSCheck = $true
    Code2016
}

# If Windows 2012, 2012 R2 or 2016 are found, exit with error
if ($OSCheck -ne $true) {
    Write-Host " "
    Write-Host "The server is not running Windows 2012, 2012 R2 or 2016.  Exiting the script."  -foregroundcolor Red
    Write-Host " "
    Exit
}