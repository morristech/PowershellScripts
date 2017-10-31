#####################################################################################################################################################################
#  SCRIPT DETAILS																																					#
#	Configures the necessary prerequisites to install Exchange 2013 CU7+ on a Windows Server 2012 (R2) server.											            #
#	Installs all required Windows 2012 (R2) components and configures service startup settings. Provides options for downloading latest Update Rollup	            #
#	and more.  First the script will determine the version of the OS you are running and then provide the correct menu items. 										#
#																																									#
# SCRIPT VERSION HISTORY																																			#
#	Current Version		: 1.17  																																	#
#	Change Log			: 1.17 - Fixed .NET 4.6.1 installation, added .NET 4.6.1 and corrected the RC4 registry entry to DWORD per Microsoft						#
#						:			Removed Windows 2008 R2 support    																			                    #
#                       : 1.16 -                                                                                                                                    #
#                       : 1.15 - Tweaked a couple of directories, corrected bad code and further testing, added PageFile Configuration, reduced length to 1400 lines#
#						: 1.14 - Completely recoded, removed duplicate code, removed old code, added colors to menu													#
#    					: 1.13 - Added hotfix for .NET 4.6.1 and changed the menu for .4.6.1 and .4.5.1 (2012 (R2) Only)											#
#                       : 1.12 - Added .NET 4.6.1 for both 2008 and 2012(R2) - good ONLY for CU13 +, removed some old code											#
#                       : 1.11 - Tweak Office Filter Pack and C++ installation and removal																			#
#                       : 1.10 - Bug Fixes																															#
#                       : 1.9 - Added a way to disable SSL 3.0 and RC4 encryption.																					#
#				        : 1.8 - Added PowerManagement																												#
#				        : 1.7 - Removed old versions of .NET (performance isue) and Windows Framework 3.0, add Edge Transport chk and Office 2010 SP2 Filter Pack	#
#				        : 1.6 - Added support for Windows 2012 R2, added options for Edge Role installation and cleaned up old items								#
#				        : 1.5 - Added support for Exchange 2013 RTM CU1, additional error suppression																#
#				        : 1.4 - Added support for Exchange 2013 RTM																									#
#				        : 1.3 - Fixed Reboot for Windows Server 2012 RTM																							#
#				        : 1.2 - fixed install commands for Windows Server 2012.  Split CAS/MX role install.															#
#				        : 1.1 - Added Windows Server 2012 Preview support																							#
#				        : 1.0 - Created script for Windows Server 2008 R2 installs																					#
#																																									#
# OTHER SCRIPT INFORMATION																																			#
#    Wish list			: Better comment based help																													#
#						: static port mapping																														#
#				        : event log logging																															#
#    Rights Required	: Local admin on server																														#
#    Exchange Version	: 2013																																		#
#    Author       		: Damian Scoles																																#
#    My Blog			: http://justaucguy.wordpress.com																											#
#    Disclaimer   		: You are on your own.  This was not written by, supported by, or endorsed by Microsoft.													#
#    Info Stolen from 	: Anderson Patricio, Bhargav Shukla and Pat Richard [Exchange 2010 script]																	#
#    					: http://msmvps.com/blogs/andersonpatricio/archive/2009/11/13/installing-exchange-server-2010-pre-requisites-on-windows-server-2008-r2.aspx #
#						: http://www.bhargavs.com/index.php/powershell/2009/11/script-to-install-exchange-2010-pre-requisites-for-windows-server-2008-r2/			#
# 						: SQL Soldier - http://www.sqlsoldier.com/wp/sqlserver/enabling-high-performance-power-plan-via-powershell									#
#																																									#
# EXECUTION																																							#
#	.\Set-Exchange2013Prerequisites-1-17.ps1																														#
#																																									#
#####################################################################################################################################################################

##################################
#   Global 	 Definitions  #
##################################
$ver = (Get-WMIObject win32_OperatingSystem).Version
$UCMAHold = $false
$OSCheck = $false
$Choice = "None"
$Date = get-date -Format "MM.dd.yyyy-hh.mm-tt"
$DownloadFolder = "c:\install"
$currentpath = (Get-Item -Path ".\" -Verbose).FullName
$Reboot = $false
$error.clear()
Start-Transcript -path "$CurrentPath\$date-Set-Prerequisites.txt" | Out-Null
Clear-Host
# Pushd

############################################################
#   Global Functions - 2012 (R2)   #
############################################################

# Begin FileDownload function
function FileDownload {
    param ($file)
    $Internetaccess = (Get-NetConnectionProfile).ipv4connectivity
    If ($Internetaccess -eq "Internet") {
        if (Test-path $DownloadFolder) {
            write-host "Target folder $DownloadFolder exists." -ForegroundColor white
        } else {
            New-Item $DownloadFolder -type Directory | Out-Null
        }
        
        BITSCheck
        
        # [string] $DownloadFile = $sourcefile.Substring($sourcefile.LastIndexOf("/") + 1)
        $DownloadFile = $File.Split([char]0x02F)
        $DownloadedFile = $DownloadFile[-1]

        if (Test-Path "$DownloadFolder\$DownloadedFile"){
            write-host "The file $DownloadFile already exists in the $DownloadFolder folder." -ForegroundColor Cyan
        } else {
            Start-BitsTransfer -Source "$File" -Destination “$DownloadFolder\$DownloadedFile"
        }
    } else {
        write-host "This machine does not have internet access and thus cannot download required files. Please resolve!" -ForegroundColor Red
    }
} # End FileDownload function

# Begin BITSCheck function
function BITSCheck {
    $Bits = Get-Module BitsTransfer
    if ($Bits -eq $null) {
        write-host "Importing the BITS module." -ForegroundColor cyan
        try {
            Import-Module BitsTransfer -erroraction STOP
        } catch {
            write-host "Server Management module could not be loaded." -ForegroundColor Red
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
            write-host " ";write-host "Server Manager module could not be loaded." -ForegroundColor Red
        }
    } else {
        # write-host "Server Manager module is already imported." -ForegroundColor Cyan
        # return $null
    }
    write-host " "
} # End ModuleStatus function

# Configure the Server for the High Performance power plan
function highperformance {
    write-host " "
	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -ne $HighPerf) {
		powercfg -setactive $HighPerf
		CheckPowerPlan
	} else {
		if ($CurrPlan -eq $HighPerf) {
			write-host " ";write-host "The power plan is already set to " -nonewline;write-host "High Performance." -ForegroundColor green;write-host " "
		}
	}
}

# Check the server power management
function CheckPowerPlan {
	$HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
	$CurrPlan = $(powercfg -getactivescheme).split()[3]
	if ($CurrPlan -eq $HighPerf) {
		write-host " ";write-host "The power plan now is set to " -nonewline;write-host "High Performance." -ForegroundColor green
	}
}

# Turn off NIC power management
function PowerMgmt {
    write-host " "
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
			If($PnPCapabilities -eq 24) {write-host " ";write-host "Power Management has already been " -NoNewline;write-host "disabled" -ForegroundColor Green;write-host " "}
   		 } 
 	 } 
 }

 # Disable RC4
function DisableRC4 {
    write-host " "
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
	write-host " "

# Add missing registry keys as needed
	If ($base -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL", $true)
		$key.CreateSubKey('Ciphers')
		$key.Close()
	} else {
		write-host "The " -nonewline
        write-host "Ciphers" -ForegroundColor green -NoNewline
        write-host " Registry key already exists."
	}

	If ($val1 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 128/128')
		$key.Close()
	} else {
		write-host "The " -nonewline
        write-host "Ciphers\RC4 128/128" -ForegroundColor green -NoNewline
        write-host " Registry key already exists."
	}

	If ($val2 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 40/128')
		$key.Close()
		New-ItemProperty -Path $registryPath2 -Name $name -Value $value -force -PropertyType "DWord"
	} else {
		write-host "The " -nonewline
        write-host "Ciphers\RC4 40/128" -ForegroundColor green -NoNewline
        write-host " Registry key already exists."
	}

	If ($val3 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers", $true)
		$key.CreateSubKey('RC4 56/128')
		$key.Close()
	} else {
		write-host "The " -nonewline
        write-host "Ciphers\RC4 56/128" -ForegroundColor green -NoNewline
        write-host " Registry key already exists."
	}
	
# Add the enabled value to disable RC4 Encryption
	If ($checkval1.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath1 -Name $name -Value $value -force -PropertyType "DWord"
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline
        write-host "Enabled" -ForegroundColor green -NoNewline
        write-host " exists under the RC4 128/128 Registry Key."
        $ssl++
	}
	If ($checkval2.enabled -ne "0") {
		write-host $checkval2
		try {
			New-ItemProperty -Path $registryPath2 -Name $name -Value $value -force -PropertyType "DWord"
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline
        write-host "Enabled" -ForegroundColor green -NoNewline
        write-host " exists under the RC4 40/128 Registry Key."
        $ssl++
	}
	If ($checkval3.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath3 -Name $name -Value $value -force -PropertyType "DWord"
            $ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline
        write-host "Enabled" -ForegroundColor green -NoNewline
        write-host " exists under the RC4 56/128 Registry Key."
        $ssl++
	}

# SSL Check totals
	If ($ssl -eq "3") {
		write-host " "
        write-host "RC4 " -ForegroundColor yellow -NoNewline
        write-host "is completely disabled on this server."
        write-host " "
	} 
	If ($ssl -lt "3"){
		write-host " "
        write-host "RC4 " -ForegroundColor yellow -NoNewline
        write-host "only has $ssl part(s) of 3 disabled.  Please check the registry to manually to add these values"
        write-host " "
	}
} # End of Disable RC4 function

# Disable SSL 3.0
function DisableSSL3 {
    write-host " "
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
		write-host "The " -nonewline;write-host "SSL 3.0" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

# Check for SSL 3.0\Server Reg Key
	If ($TestPath2 -eq $null) {
		$key = (get-item HKLM:\).OpenSubKey("System\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0", $true)
		$key.CreateSubKey('Server')
		$key.Close()
	} else {
		write-host "The " -nonewline;write-host "SSL 3.0\Servers" -ForegroundColor green -NoNewline;write-host " Registry key already exists."
	}

# Add the enabled value to disable SSL 3.0 Support
	If ($checkval1.enabled -ne "0") {
		try {
			New-ItemProperty -Path $registryPath -Name $name -Value $value -force;$ssl++
		} catch {
			$SSL--
		} 
	} else {
		write-host "The registry value " -nonewline;write-host "Enabled" -ForegroundColor green -NoNewline;write-host " exists under the SSL 3.0\Server Registry Key."
	}
} # End of Disable SSL 3.0 function

# Function - Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-WinUniComm4 {
    write-host " "
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		if($val.DisplayVersion -ne "5.0.8132.0"){
			if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is not installed.  Downloading and installing now." -ForegroundColor yellow
				Install-NewWinUniComm4
			} else {
    				Write-Host "`nAn old version of Microsoft Unified Communications Managed API 4.0 is installed."
				UnInstall-WinUniComm4
				Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now."  -ForegroundColor green
				Install-NewWinUniComm4
			}
   		} else {
   			Write-Host "`nThe Preview version of Microsoft Unified Communications Managed API 4.0 is installed."
   			UnInstall-WinUniComm4
   			Write-Host "`nMicrosoft Unified Communications Managed API 4.0 has been uninstalled.  Downloading and installing now." -ForegroundColor green
   			Install-NewWinUniComm4
		}
	} else {
		write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
		write-host "installed." -ForegroundColor green
	}
    write-host " "
} # end Install-WinUniComm4

# Install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function Install-NewWinUniComm4{
	$file = "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
    FileDownload $File
	Set-Location $DownloadFolder
    # [string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $downloadfolder\WinUniComm4.log"
	Write-Host "File: UcmaRuntimeSetup.exe installing..." -NoNewLine
	# Invoke-Expression $expression

    Start-Process '.\NDP461-KB3102436-x86-x64-AllOS-ENU.exe' -ArgumentList '/quiet','/norestart' –Wait 

	Start-Sleep -Seconds 20
	$val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
	if($val.DisplayVersion -ne "5.0.8308.0"){
		Write-Host "`nMicrosoft Unified Communications Managed API 4.0 is now installed" -ForegroundColor Green
	}
    write-host " "
} # end Install-NewWinUniComm4

# Uninstall Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
function UnInstall-WinUniComm4{
	FileDownload "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe"
 	Set-Location $DownloadFolder
  	[string]$expression = ".\UcmaRuntimeSetup.exe /quiet /norestart /l* $downloadfolder\WinUniComm4.log"
  	Write-Host "File: UcmaRuntimeSetup.exe uninstalling..." -NoNewLine
   	Invoke-Expression $expression
  	Start-Sleep -Seconds 20
	if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}") -eq $false){
		write-host "Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
		write-host "been uninstalled!" -ForegroundColor red
	}
} # end Uninstall-WinUniComm4

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

# Function - Windows Management Framework 4.0 - Install - Needed for CU3+
function Install-WinMgmtFW4{
    # Windows Management Framework 4.0
	$wmf = $PSVersionTable.psversion
	if ($wmf.major -eq "4") {
	    	Write-Host "`nWindows Management Framework 4.0 is already installed" -ForegroundColor Green
	} else {
	    	FileDownload "http://download.microsoft.com/download/3/D/6/3D61D262-8549-4769-A660-230B67E15B25/Windows8-RT-KB2799888-x64.msu"
    		Set-Location $DownloadFolder
	    	[string]$expression = ".\Windows8-RT-KB2799888-x64.msu /quiet /norestart"
	    	Write-Host "File: Windows8-RT-KB2799888-x64 installing..." -NoNewLine
	    	Invoke-Expression $expression
    		Start-Sleep -Seconds 20
		$wmf = $PSVersionTable.psversion
	
	    	if ($wmf.major -ge "4") {Write-Host "`b`b`b`b`b`b`b`b`b`b`b`b`binstalled!   " -ForegroundColor Green} else {Write-Host "`b`b`b`b`b`b`b`b`b`b`b`b`bFAILED!" -ForegroundColor Red}
    }
} # end Install-WinMgmtFW4

# Final Cleanup - C++ and register ASP .NET
function Cleanup-Final {
	# Old C++ from the old UCMA
	# [STRING] $downloadfile2 = "C:\ProgramData\Package Cache\{5b2d190f-406e-49cf-8fea-1c3fc6777778}"
	[STRING] $downloadfile2 = "C:\ProgramData\Package Cache\{15134cb0-b767-4960-a911-f2d16ae54797}"
	Set-Location $DownloadFolder2
	[string]$expression = ".\vcredist_x64.exe /q /uninstall /norestart"
	Invoke-Expression $expression
	c:\Windows\Microsoft.NET\Framework64\v4.0.30319\aspnet_regiis.exe -ir -enable
	iisreset
}

# Function - Windows Management Framework 4.0 - Install - Needed for CU3+
function Install-WinMgmtFW4{
    # Windows Management Framework 4.0
	$wmf = $PSVersionTable.psversion
	if ($wmf.major -eq "4") {
	    	Write-Host "`nWindows Management Framework 4.0 is already installed" -ForegroundColor Green
	} else {
	    FileDownload "http://download.microsoft.com/download/3/D/6/3D61D262-8549-4769-A660-230B67E15B25/Windows8-RT-KB2799888-x64.msu"
    	Set-Location $DownloadFolder
	    [string]$expression = ".\Windows8-RT-KB2799888-x64.msu /quiet /norestart"
	    Write-Host "File: Windows8-RT-KB2799888-x64 installing..." -NoNewLine
	    Invoke-Expression $expression
    	Start-Sleep -Seconds 20
		$wmf = $PSVersionTable.psversion
	    if ($wmf.major -ge "4") {Write-Host "`b`b`b`b`b`b`b`b`b`b`b`b`binstalled!   " -ForegroundColor Green} else {Write-Host "`b`b`b`b`b`b`b`b`b`b`b`b`bFAILED!" -ForegroundColor Red}
    }
} # end Install-WinMgmtFW4

# Mailbox requirements - Part 1
function check-prereqset1 {

    # .NET 4.5.2, 4.6.1 or 4.6.1
	Check-DotNetVersion

    # Windows Management Framework 4.0 - Check - Needed for CU3+
	$wmf = $PSVersionTable.psversion
	if ($wmf.major -ge "4") {
        Write-Host "Windows Management Framework 4.0 is " -nonewline 
	    write-host "installed." -ForegroundColor green
	} else {
	    write-host "Windows Management Framework 4.0 is " -nonewline 
	    write-host "not installed!" -ForegroundColor red
	}

    # Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit 
    $val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{41D635FE-4F9D-47F7-8230-9B29D6D42D31}" -Name "DisplayVersion" -erroraction silentlycontinue
    if($val.DisplayVersion -ne "5.0.8308.0"){
        if($val.DisplayVersion -ne "5.0.8132.0"){
            if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A41CBE7D-949C-41DD-9869-ABBD99D753DA}") -eq $false) {
                write-host "No version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
                write-host "not installed!" -ForegroundColor red
                write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
            }else {
            write-host "The Preview version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline 
            write-host "installed." -ForegroundColor red
            write-host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red
            write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
            }
        } else {
        write-host "The wrong version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
        write-host "installed." -ForegroundColor red
        write-host "This is the incorrect version of UCMA. "  -nonewline -ForegroundColor red 
        write-host "Please install the newest UCMA 4.0 from http://www.microsoft.com/en-us/download/details.aspx?id=34992." 
        }   
    } else {
         write-host "The correct version of Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit is " -nonewline
         write-host "installed." -ForegroundColor green
    }
} # End mailbox requirements - Part 1

# CAS requirements - Part 1 / Mailbox requirements - Part 2
function check-prereqset2 {

    # Windows Identity Foundation
	$hotfix1 = Get-HotFix -id KB974405 -ErrorAction SilentlyContinue
    if ($hotfix1 -match "KB974405") {
	    Write-Host "Windows Identity Foundation is " -nonewline 
	    write-host "installed." -ForegroundColor green
    } else {
	    Write-Host "Windows Identity Foundation is " -nonewline 
	    write-host "not installed!" -ForegroundColor red
	}

    # Association Cookie/GUID used by RPC over HTTP Hotfix
	$hotfix1 = Get-HotFix -id KB2619234 -ErrorAction SilentlyContinue
	if ($hotfix1 -match "KB2619234") {
    	Write-Host "Association Cookie/GUID used by RPC over HTTP Hotfix is " -nonewline 
	    write-host "installed." -ForegroundColor green
    } else {
	    Write-Host "`nAssociation Cookie/GUID used by RPC over HTTP Hotfix is " -nonewline 
	    write-host "not installed!" -ForegroundColor red
	}

    # Insecure library loading could allow remote code execution
	$hotfix1 = Get-HotFix -id KB2533623 -ErrorAction SilentlyContinue
	if ($hotfix1 -match "KB2533623") {
    	Write-Host "Insecure library loading could allow remote code execution is " -nonewline 
	    write-host "installed." -ForegroundColor green
    } else {
	    Write-Host "Insecure library loading could allow remote code execution is " -nonewline 
	    write-host "not installed!" -ForegroundColor red
	}

    # Check for C++ Install and ASP .NEt
    $directory = get-item "C:\ProgramData\Package Cache\{15134cb0-b767-4960-a911-f2d16ae54797}" -ErrorAction SilentlyContinue
    if ($directory -match "{15134cb0-b767-4960-a911-f2d16ae54797}") {
		write-host "Microsoft Visual C++ has " -nonewline
		write-host "not been uninstalled!" -ForegroundColor red
	} else {
		write-host "Microsoft Visual C++ has been " -nonewline
		write-host "uninstalled!" -ForegroundColor green
	}

    # Exist function
	write-host "Make sure you registered ASP .Net as well.  See here " -nonewline
	write-host "http://technet.microsoft.com/en-us/library/bb691354(v=exchg.150).aspx" -ForegroundColor yellow
} # End CAS Req. - Part 1 / Mailbox requirements - Part 2

# Check Windows features on 2012 CAS Servers
function Check-winfeaturesCAS2012 {
	$values = @("AS-HTTP-Activation","Desktop-Experience","NET-Framework-45-Features","RPC-over-HTTP-proxy","RSAT-Clustering","Web-Mgmt-Console","WAS-Process-Model","Web-Asp-Net45","Web-Basic-Auth","Web-Client-Auth","Web-Digest-Auth","Web-Dir-Browsing","Web-Dyn-Compression","Web-Http-Errors","Web-Http-Logging","Web-Http-Redirect","Web-Http-Tracing","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Lgcy-Mgmt-Console","Web-Metabase","Web-Mgmt-Console","Web-Mgmt-Service","Web-Net-Ext45","Web-Request-Monitor","Web-Server","Web-Stat-Compression","Web-Static-Content","Web-Windows-Auth","Web-WMI","Windows-Identity-Foundation")
	foreach ($item in $values){
	    $val = get-Windowsfeature $item
	    If ($val.installed -eq $true){
	        write-host "The Windows Feature"$item" is " -nonewline 
	        write-host "installed." -ForegroundColor green
	    } else {
	        write-host "The Windows Feature"$item" is " -nonewline 
	        write-host "not installed!" -ForegroundColor red
	    }
	}
}

# Check Windows features on 2012 MBX Servers
function Check-winfeaturesMBX2012 {
	$values = @("AS-HTTP-Activation","Desktop-Experience","NET-Framework-45-Features","RPC-over-HTTP-proxy","RSAT-Clustering","RSAT-Clustering-CmdInterface","Web-Mgmt-Console","WAS-Process-Model","Web-Asp-Net45","Web-Basic-Auth","Web-Client-Auth","Web-Digest-Auth","Web-Dir-Browsing","Web-Dyn-Compression","Web-Http-Errors","Web-Http-Logging","Web-Http-Redirect","Web-Http-Tracing","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Lgcy-Mgmt-Console","Web-Metabase","Web-Mgmt-Console","Web-Mgmt-Service","Web-Net-Ext45","Web-Request-Monitor","Web-Server","Web-Stat-Compression","Web-Static-Content","Web-Windows-Auth","Web-WMI","Windows-Identity-Foundation")
	foreach ($item in $values){
	    $val = get-Windowsfeature $item
	    If ($val.installed -eq $true){
	        write-host "The Windows Feature"$item" is " -nonewline 
	        write-host "installed." -ForegroundColor green
	    } else {
	        write-host "The Windows Feature"$item" is " -nonewline 
	        write-host "not installed!" -ForegroundColor red
	    }
	}
}

# Edge Transport requirements
function check-prereqset5 {
    # Windows Feature AD LightWeight Services
	$values = @("ADLDS")
	foreach ($item in $values){
		$val = get-Windowsfeature $item
		If ($val.installed -eq $true){
			write-host "The Windows Feature"$item" is " -nonewline 
			write-host "installed." -ForegroundColor green
		}else{
			write-host "The Windows Feature"$item" is " -nonewline 
			write-host "not installed!" -ForegroundColor red
		}
	}

    # .NET 4.5.2 [for CU7+] or .NET 4.6.1 [CU13+]
    Check-DotNetVersion

    # Windows Management Framework 4.0 - Check - Needed for CU3+
	$wmf = $PSVersionTable.psversion
	if ($wmf.major -ge "4") {
    	Write-Host "Windows Management Framework 4.0 is " -nonewline 
	    write-host "installed." -ForegroundColor green
	} else {
	    write-host "Windows Management Framework 4.0 is " -nonewline 
	    write-host "not installed!" -ForegroundColor red
	}
} # End Edge Transport requirements

# CAS Requirements - Part 3
function check-prereqset6 {
    write-host ""
	write-host "Make sure to open port 139 in the Windows firewall:"
	write-host "http://technet.microsoft.com/en-us/library/bb691354(v=exchg.150).aspx" -ForegroundColor yellow
} # End CAS Requirements - Part 3

# Configure Net TCP Port Sharing - RunOnce
function NetTCPPortSharing {

	$Server = (hostname)
	$NetTCP = "Set-Content \\$server config NetTcpPortSharing start= auto"
	if (Get-ItemProperty -Name "NetTCPPortSharing" -path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce' -ErrorAction SilentlyContinue) { 
	    Write-host "Registry key HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce\NetTCPPortSharing already exists." -ForegroundColor yellow
		Set-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "NetTCPPortSharing" -Value $NetTCP | Out-Null
	} else { 
	    New-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "NetTCPPortSharing" -Value $NetTCP -PropertyType "String" | Out-Null
	} 

} # End configure Net TCP Port Sharing

# Configure PageFile for Exchange
function ConfigurePagefile {
    $stop = $false

    # Remove Existing PageFile
    try {
        Set-CimInstance -Query “Select * from win32_computersystem” -Property @{automaticmanagedpagefile=”False”} 
    } catch {
        write-host "Cannot remove the existing Pagefile." -ForegroundColor Red
        $stop = $true
    }
    # Get RAM and set ideal PageFileSize
    $GB = 1048576
    try {
        $RamInMb = (Get-CIMInstance -computername $name -Classname win32_physicalmemory -ErrorAction Stop | measure-object -property capacity -sum).sum/$GB
    } catch {
        write-host "Cannot acquire the amount of RAM in the server." -ForegroundColor Red
        $stop = $true
    }
    $ExchangeRAM = $RAMinMb + 10

    if ($stop -ne $true) {
        # Configure PageFile
        try {
            Set-CimInstance -Query “Select * from win32_PageFileSetting” -Property @{InitialSize=$ExchangeRAM;MaximumSize=$ExchangeRAM}
        } catch {
            write-host "Cannot configure the PageFile correctly." -ForegroundColor Red
        }
        $pagefile = Get-CimInstance win32_PageFileSetting -Property * | select-object Name,initialsize,maximumsize
        $name = $pagefile.name;$max = $pagefile.maximumsize;$min = $pagefile.initialsize
        write-host " "
        write-host "The page file of $name is now configured for an initial size of " -ForegroundColor white -NoNewline
        write-host "$min " -ForegroundColor green -NoNewline
        write-host "and a maximum size of " -ForegroundColor white -NoNewline
        write-host "$max." -ForegroundColor Green
        write-host " "
    } else {
        write-host "The PageFile cannot be configured at this time." -ForegroundColor Red
    }
}

######################################################
#    This section is for the Windows 2012 (R2) OS    #
######################################################`

function Code2012 {

# Start code block for Windows 2012 or 2012 R2
$Menu2012 = {
	write-host "	******************************************************************" -ForegroundColor cyan
	write-host "	Exchange Server 2013 [On Windows 2012 / 2012 R2] - Features script" -ForegroundColor cyan
	write-host "	******************************************************************" -ForegroundColor cyan
	write-host "	"
	write-host "	** .NET 4.5.2 installation has been removed **" -ForegroundColor Red
	write-host "	"
	write-host "	Please select an option from the list below." -ForegroundColor yellow
    write-host "	"
	write-host "	Exchange Server 2013 CU14 [.NET 4.6.2]" -ForegroundColor yellow
    write-host "	1) Install Client Access Server prerequisites - Step 1 [Includes 30 & 31] " -ForegroundColor white
	write-host "	2) Install Client Access Server prerequisites - Step 2" -ForegroundColor white
	write-host "	3) Install Mailbox and or CAS/Mailbox prerequisites - Step 1 [Includes 30 & 31]" -ForegroundColor white
	write-host "	4) Install Mailbox and or CAS/Mailbox prerequisites - Step 2" -ForegroundColor white
	write-host "	5) Install Edge Transport Server prerequisites" -ForegroundColor white
    write-host "	"
	write-host "	Exchange Server 2013 CU13 and 14 [.NET 4.6.1]" -ForegroundColor yellow
	write-host "	6) Install Client Access Server prerequisites - Step 1 [Includes 30 & 31] " -ForegroundColor white
	write-host "	7) Install Client Access Server prerequisites - Step 2" -ForegroundColor white
	write-host "	8) Install Mailbox and or CAS/Mailbox prerequisites - Step 1 [Includes 30 & 31]" -ForegroundColor white
	write-host "	9) Install Mailbox and or CAS/Mailbox prerequisites - Step 2" -ForegroundColor white
	write-host "	10) Install Edge Transport Server prerequisites - Step 1" -ForegroundColor white
	write-host "	11) Install Edge Transport Server prerequisites - Step 2" -ForegroundColor white
    write-host "	"
	write-host "	Miscellaneous" -ForegroundColor yellow
	write-host "	15) Launch Windows Update" -ForegroundColor white
	write-host "	16) Check Prerequisites for CAS role" -ForegroundColor white
	write-host "	17) Check Prerequisites for Mailbox role or Cas/Mailbox roles" -ForegroundColor white
	write-host "	18) Check Prerequisites for Edge role" -ForegroundColor white
    write-host "	"
	write-host "	One Off Changes" -ForegroundColor yellow
	write-host "	20) Install - One Off - STEP 1 - Windows Components - CAS role" -ForegroundColor white
	write-host "	21) Install - One Off - STEP 1 - Windows Components - Mailbox (or CAS/Mailbox) Role" -ForegroundColor white
	write-host "	22) Install - One Off - STEP 4 - Unified Communications Managed API 4.0" -ForegroundColor white
	write-host "	23) Install - One Off - .NET 4.6.2" -ForegroundColor white
	write-host "	24) Install - One Off - .NET 4.6.1 - Prerequisites" -ForegroundColor white
	write-host "	25) Install - One Off - .NET 4.6.1 and post hotfixes" -ForegroundColor white
    write-host "	"
	write-host "	New Features" -ForegroundColor yellow
	write-host "	30) Set Power Plan to High Performance" -ForegroundColor white
	write-host "	31) Disable Power Management for NICs" -ForegroundColor white
	write-host "	32) Disable SSL 3.0 Support" -ForegroundColor white
	write-host "	33) Disable RC4 Support" -ForegroundColor white
    write-host "	"
    write-host "	POST EXCHANGE 2013 INSTALL" -ForegroundColor Yellow
    write-host "	40) Configure PageFile to RAM + 10 MB" -ForegroundColor green
    write-host "		"
	write-host "	98) Restart the Server" -ForegroundColor red
	write-host "	99) Exit" -ForegroundColor cyan
    write-host "	"
    write-host "	Select an option.. [1-99]? " -ForegroundColor white -nonewline
}

################################
#        2012 Functions        #
################################

# Add a firewall rule for CAS role - Port 
function Add-FirewallRule {
   param( 
      $name,
      $tcpPorts,
      $appName = $null,
      $serviceName = $null
   )
    $fw = New-Object -ComObject hnetcfg.fwpolicy2 
    $rule = New-Object -ComObject HNetCfg.FWRule
        
    $rule.Name = $name
    if ($appName -ne $null) { $rule.ApplicationName = $appName }
    if ($serviceName -ne $null) { $rule.serviceName = $serviceName }
    $rule.Protocol = 6 #NET_FW_IP_PROTOCOL_TCP
    $rule.LocalPorts = $tcpPorts
    $rule.Enabled = $true
    $rule.Grouping = "@firewallapi.dll,-23255"
    $rule.Profiles = 7 # all
    $rule.Action = 1 # NET_FW_ACTION_ALLOW
    $rule.EdgeTraversal = $false
    
    $fw.Rules.Add($rule)
}

################################
#     2012 Menu Backend        #
################################

    Do { 	
	    if ($Reboot -eq $true){Write-Host "REBOOT REQUIRED!" -backgroundcolor red -ForegroundColor black;Write-Host "DO NOT INSTALL EXCHANGE BEFORE REBOOTING!" -backgroundcolor red -ForegroundColor black}
	    if ($Choice -ne "None") {Write-Host "Last command: "$Choice -ForegroundColor Yellow}	
        invoke-command -scriptblock $Menu2012
	    $Choice = Read-Host

        switch ($Choice)    {

            # -- 4.6.2 --

            1 {# 	Prep CAS - Step 1
                ModuleStatus
			    NetTCPPortSharing
			    HighPerformance
			    PowerMgmt
			    Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation        
			    Add-FirewallRule "Exchange Server 2013 - CAS" "139" $null $null
			    $RebootRequired = $true
		    }
		    2 {#	Prep CAS - Step 2
                ModuleStatus
			    Install-DotNET462
                Install-WinUniComm4
                Install-WinMgmtFW4
			    $RebootRequired = $true
		    }
		    3 {# 	Prep Mailbox or CAS/Mailbox - Step 1
                ModuleStatus
			    NetTCPPortSharing
			    HighPerformance
			    PowerMgmt
			    Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
			    $RebootRequired = $true
		    }
		    4 {#	Prep Mailbox or CAS/Mailbox - Step 2
                ModuleStatus
                Install-DotNET462
                Install-WinUniComm4
                Install-WinMgmtFW4
			    $RebootRequired = $true
		    }
	  	    5 {#	Prep Exchange Transport
			    Install-windowsfeature ADLDS
			    Install-DotNET462
			    Install-WinMgmtFW4
		    }
		
            # -- 4.6.1 --

            6 {# 	Prep CAS - Step 1
			    # Get-ModuleStatus -name ServerManager
                ModuleStatus
                Install-PreNET461
			    NetTCPPortSharing
			    HighPerformance
			    PowerMgmt
			    Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation        
			    Add-FirewallRule "Exchange Server 2013 - CAS" "139" $null $null
			    $RebootRequired = $true
		    }
		    7 {#	Prep CAS - Step 2
			    # Get-ModuleStatus -name ServerManager
                ModuleStatus
			    Install-DotNET461
                POSTDotNET461
                Install-WinUniComm4
                Install-WinMgmtFW4
			    $RebootRequired = $true
		    }
		    8 {# 	Prep Mailbox or CAS/Mailbox - Step 1
			    # Get-ModuleStatus -name ServerManager
                ModuleStatus
                Install-PreNET461
			    NetTCPPortSharing
			    HighPerformance
			    PowerMgmt
			    Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
			    $RebootRequired = $true
		    }
		    9 {#	Prep Mailbox or CAS/Mailbox - Step 2
			    Install-DotNET461
                POSTDotNET461
                Install-WinUniComm4
                Install-WinMgmtFW4
			    $RebootRequired = $true
		    }
	  	    10 {#	Prep Exchange Transport - Step 1
			    Install-windowsfeature ADLDS
			    Install-PreNET461
		    }
	  	    11 {#	Prep Exchange Transport - Step 2
			    Install-DotNET461
                POSTDotNET461
			    Install-WinMgmtFW4
		    }

            # -- Misc 1 --

	  	    15 {#	Windows Update
			    Invoke-Expression "$env:windir\system32\wuapp.exe startmenu"
		    }
		    16 {# 	CAS Requirement Check
			    check-prereqset1
			    # check-prereqset4
                check-prereqset6
                Check-winfeaturesCAS2012
		    }
		    17 {#	Mailbox or CAS/Mailbox Requirement Check
			    check-prereqset1
                # check-prereqset3
                Check-winfeaturesCAS2012
		    }
		    18 {#	Edge Transport Requirement Check
			    check-prereqset5
		    }

            # -- One Off Changes --

		    20 {#	Step 1 - One Off - Windows Components - CAS
			    Get-ModuleStatus -name ServerManager
			    Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
		    }
		    21 {#	Step 1 - One Off - Windows Components - Mailbox or CAs/Mailbox
			    Get-ModuleStatus -name ServerManager
			    Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation
		    }
		    22 {#	Install - One Off - Unified Communications Managed API 4.0
			    Install-WinUniComm4
		    }
		    23 {#	Install - One Off - .NET 4.6.2
			    Install-DotNET462
		    }
		    24 {#	Install - One Off - .NET 4.6.1 Prerequisites
			    Install-PreNET461
		    }
		    25 {#	Install - One Off - .NET 4.6.1 and post hotfix
			    Install-DotNET461
                POSTDotNET461
		    }

            # -- New Feautures --

		    30 { # Set power plan to High Performance as per Microsoft
			    highperformance
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
            40 {#   Configure the pagefile to be RAM + 10 and not system managed
                ConfigurePageFile
            }
		    98 {#	Exit and restart
			    # Stop-Transcript
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
			    # Stop-Transcript
		    }
		    default {Write-Host "You haven't selected any of the available options. "}
	    }
    } while ($Choice -ne 99)
}

######################################################
#               MAIN SCRIPT BODY                     #
######################################################

# Check for Windows 2012 or Windows 2012 R2
if (($ver -match '6.2') -or ($ver -match '6.3')) {
    $OSCheck = $true
    Code2012
}

# If Windows 2008, 2012 or 2012 R2 are found, exit with error
if ($OSCheck -ne $true) {
    write-host " "
    write-host "The server is not running Windows 2012 or 2012 R2.  Exiting the script."  -ForegroundColor Red
    write-host " "
    Exit
}