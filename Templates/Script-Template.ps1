<#
    .NOTES
    --------------------------------------------------------------------------------
     Created on:       DD/MM/YYY
     Created by:       James Stapleton
     Version:          0.0.0.1

     Modifications:    
    --------------------------------------------------------------------------------
#>

#region - Main Program Try
#---------------------------------------------------------[Main Program Try]--------------------------------------------------------
Try # Start of Main Program Try
{
	#---------------------------------------------------------[Initialisations]--------------------------------------------------------
	$global:LogFilePath = "c:\temp\testingLog.log" # Log file path
	$ErrorActionPreference = 'Stop' # Set Error Action
	Set-Location $PSScriptRoot # Set directory location to script route
	
	$scriptVersion = "0.0.0.1"
		
	#-----------------------------------------------------------[Functions]------------------------------------------------------------
	#region - Function: Main
	Function Main {
		
		Write-Log "*** Function Main Started ***"
		
	} # Main: End
	#endregion
	
	#region - Function: Write-Log
	Function Write-Log {
		[CmdletBinding()]
		param (
			[Parameter(
					   Mandatory = $true)]
			[string]$Message,
			[Parameter()]
			[ValidateSet(1, 2, 3)]
			[int]$LogLevel = 1,
			[int]$LogCount = 0
		)
		DO {
			If (Test-Path $global:LogFilePath) {
				$TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
				$Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
				$LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)", $LogLevel
				$Line = $Line -f $LineFormat
				
				If ($LogLevel -eq 1) { Write-Output $Message }
				else { Write-Warning -Message $Message }
				Add-Content -Value $Line -Path $global:LogFilePath
				$LogCount = 1
			}
			else {
				If ($LogCount -eq 0) {
					New-Item $global:LogFilePath -Type File | Out-Null
					If (-not (Test-Path $global:LogFilePath)) { $LogCount = 2 }
				}
				ElseIf ($LogCount -eq 2) { Write-Warning "Couldn't create log file: $global:LogFilePath"; Break }
			}
		}
		Until ($LogCount -eq 1)
	} # Write-Log: End
	#endregion
	
	#region - Function: Write-InitialLogEntry
	Function Write-InitialLogEntry {
		Write-Log "=================================================================="
		Write-Log "  Script Started: `t $currentDate (D/M/Y) $currentTime (local)"
		Write-Log "  Script name: $scriptName"
		Write-Log "  Script version: $scriptVersion"
		Write-Log "  Script path: $scriptPath"
		Write-Log "  This log file path: $global:LogFilePath"
		Write-Log "  Computer name: $envComputerName"
		Write-Log "  PowerShell Host Version: $($envHost.Version)"
		Write-Log "  Processor Architecture: $envArchitecture"
		Write-Log "  Is this a 64bit OS: $is64Bit"
		Write-Log "  Powershell scripting architecture: $psArchitecture"
		Write-Log "  Username: $envUserName"
		Write-Log "  Program Files directory: $envProgramFiles"
		Write-Log "  Program Files x86 directory: $envProgramFilesX86"
		Write-Log "  ProgramData directory: $envProgramData"
		Write-Log "  Windows directory: $envWinDir"
		Write-Log "  Windows Temp directory: $envWinTempDir"
		Write-Log "  Script Error Preference set as: $ErrorActionPreference"
		Write-Log "=================================================================="
		Write-Log " "
		Write-Log "Main Program start"
		Write-Log "---------------------------------"
		Write-Log " "
	}
	#endregion
	
	#region - Functions: Custom
	# Place holder to add additonal script functions if needed
	
	#endregion
	
	#----------------------------------------------------------[Declarations]----------------------------------------------------------
	#region - Variables: Environment
	$currentDate = (Get-Date -UFormat "%d-%m-%Y")
	$currentTime = (Get-Date -UFormat "%T")
	$culture = Get-Culture
	$envHost = $host
	$envAllUsersProfile = $env:ALLUSERSPROFILE
	$envAppData = $env:APPDATA
	$envArchitecture = $env:PROCESSOR_ARCHITECTURE
	$envCommonProgramFiles = $env:CommonProgramFiles
	$envCommonProgramFilesX86 = "${env:CommonProgramFiles(x86)}"
	$envComputerName = $env:COMPUTERNAME
	$envHomeDrive = $env:HOMEDRIVE
	$envHomePath = $env:HOMEPATH
	$envHomeShare = $env:HOMESHARE
	$envLocalAppData = $env:LOCALAPPDATA
	$envLogonServer = $env:LOGONSERVER
	$envOS = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
	If ($envOS.version -like "10.*") { $envOSName = "Windows10" }
	ElseIf ($envOS.version -like "6.2*") { $envOSName = "Windows8" }
	ElseIf ($envOS.version -like "6.1*") { $envOSName = "Windows7" }
	ElseIf ($envOS.version -like "6.0*") { $envOSName = "WindowsVista" }
	ElseIf ($envOS.version -like "5.2*") { $envOSName = "WindowsXPPro" }
	Else { $envOSName = "Unknown" }
	If ($envOS.version -like "10.*") {
		If (Test-Path "hkcu:\Software\Microsoft\OneDrive") { $envOneDrive = (Get-ItemProperty -Path "hkcu:\Software\Microsoft\OneDrive\" -Name UserFolder).UserFolder }
		Else { $envOneDrive = "Unknown" }
	}
	Else {
		If (Test-Path "hkcu:\Software\Microsoft\Windows\CurrentVersion\SkyDrive") { $envOneDrive = (Get-ItemProperty -Path "hkcu:\Software\Microsoft\Windows\CurrentVersion\SkyDrive\" -Name UserFolder).UserFolder }
		else { $envOneDrive = "Unknown" }
	}
	$envProgramFiles = $env:PROGRAMFILES
	$envProgramFilesX86 = "${env:ProgramFiles(x86)}"
	$envProgramData = $env:PROGRAMDATA
	$envPublic = $env:PUBLIC
	$envSystemDrive = $env:SYSTEMDRIVE
	$envSystemRoot = $env:SYSTEMROOT
	$envTemp = $env:TEMP
	$envUserDNSDomain = $env:USERDNSDOMAIN
	$envUserDomain = $env:USERDOMAIN
	$envUserName = $env:USERNAME
	$envUserProfile = $env:USERPROFILE
	$envWinDir = $env:WINDIR
	$envWinTempDir = "$envWinDir\Temp"
	# Handle X86 environment variables so they are never empty
	If ($envCommonProgramFilesX86 -eq $null -or $envCommonProgramFilesX86 -eq "") { $envCommonProgramFilesX86 = $env:CommonProgramFiles }
	If ($envProgramFilesX86 -eq $null -or $envProgramFilesX86 -eq "") { $envProgramFilesX86 = $env:PROGRAMFILES }
	$currentLanguage = $PSUICulture.SubString(0, 2).ToUpper()
	$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition)
	$scriptPath = $MyInvocation.MyCommand.Definition
	$scriptFileName = Split-Path -Leaf $MyInvocation.MyCommand.Definition
	$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$scriptCurrentDirectory = $pwd
	# Get the invoking script directory
	If (((Get-Variable MyInvocation).Value).ScriptName) {
		$scriptParentPath = Split-Path -Parent ((Get-Variable MyInvocation).Value).ScriptName
	}
	# Fall back on the directory one level above this script
	Else {
		$scriptParentPath = (Get-Item $scriptRoot).Parent.FullName
	}
	# Variables: Executables
	$exeWusa = "wusa.exe"
	$exeMsiexec = "msiexec.exe"
	$exeSchTasks = "$envWinDir\System32\schtasks.exe"
	# Variables: Architecture
	$is64Bit = (Get-WmiObject -Class Win32_OperatingSystem -ea 0).OSArchitecture -eq '64-bit'
	$is64BitProcess = [System.IntPtr]::Size -eq 8
	If ($is64BitProcess -eq $true) { $psArchitecture = "x64" }
	Else { $psArchitecture = "x86" }
	$isServerOS = (Get-WmiObject -Class Win32_operatingsystem -ErrorAction SilentlyContinue | Select Name -ExpandProperty Name)
	#endregion
	
	#-----------------------------------------------------------[Execution]------------------------------------------------------------
	
	Write-InitialLogEntry
	Main
	
} # End of Main Program Try
#endregion

#region - Main Program Catch
#---------------------------------------------------------[Main Program Catch]--------------------------------------------------------
Catch # Start of Main Program Catch
{
	# Log error messages
	$e = $_.Exception
	$line = $_.InvocationInfo.ScriptLineNumber
	$msg = $e.Message
	Write-Log "Caught exception at line `"$line`" within the following error message: $msg"
} # End of Main Program Catch
#endregion

#region - Main Program Finally
#---------------------------------------------------------[Main Program Finally]--------------------------------------------------------
Finally # Start of Main Program Catch
{
	Write-Log " "
	Write-Log "---------------------------------"
	Write-Log "Main Program End"
	Write-Log " "
	
	$currentTime = (Get-Date -UFormat "%T")
	Write-Log "=================================================================="
	Write-Log "  Script Ended: `t $currentDate (D/M/Y) $currentTime (local)"
	Write-Log "=================================================================="
	Write-Log " "
} # End of Main Program Finally
#endregion
