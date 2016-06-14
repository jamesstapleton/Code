function Remove-InstalledMSI {
  <#
  .SYNOPSIS
  Removes a given MSI based on display name provided

  .DESCRIPTION
  This function will look for a certain display name, or a display name that matches a certain string,
  and then will proceed to remove that MSI. There is extra critiera that can be used to limit what is
  removed. This is in the form of specifying a particular display name not to remove if found.

  .EXAMPLE
  Remove-InstalledMSI -AppToUninstall 'SOME_PART_OF_DISPLAYNAME' -SCCMAppName 'SL_SCCM_APPLICATION_NAME'

  .EXAMPLE
  Remove-InstalledMSI -AppToUninstall 'FULL_DISPLAYNAME_TO_REMOVE' -SCCMAppName 'SL_SCCM_APPLICATION_NAME'

  .EXAMPLE
  Remove-InstalledMSI -AppToUninstall 'SOME_PART_OF_DISPLAYNAME' -SCCMAppName 'SL_SCCM_APPLICATION_NAME' -KeepAppName 'FULL_DISPLAYNAME_TO_KEEP'

  .EXAMPLE
  From a command prompt, batch file or SCCM command
  powershell -NoLogo -NoProfile -NonInteractive -command ". .\Remove-InstalledMSI_FunctionSCCM.ps1; Remove-InstalledMSI -AppToUninstall 'SOME_PART_OF_DISPLAYNAME' -SCCMAppName 'SL_SCCM_APPLICATION_NAME' -KeepAppName ''"

  .PARAMETER AppToUninstall
  You can use part of the display name as query uses wild cards. Or you can specify the full display name.

  .PARAMETER SCCMAppName
  Give the name of the SCCM name of this uninstall application. This is used to create a txt file in order for SCCM detection if successful

  .PARAMETER KeepAppName
  Give the full name of the display name to keep if found. This is used in case you look for apps but you may want to keep the latest version.

  .NOTES
  Version 1.0.1 by James Stapleton
  Date 09/06/2016

  Credits
  Used the following information from then below link as a bases for designing this script, in that not to use a Win32_product WMI query
  https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/13/use-powershell-to-quickly-find-installed-software/
  #>

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True,
				   ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True)]
		[string[]]$AppToUninstall,
        [Parameter(Mandatory = $True,
				   ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True)]
		[string[]]$SCCMAppName,
		[Parameter(ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True)]
		[string[]]$KeepAppName
	)
	begin {
		write-verbose "Setting up variables"
        $currentDate = (Get-Date -UFormat "%d-%m-%Y")
        $currentTime = (Get-Date -UFormat "%T")
        $sccmDetectionFolder = "C:\Program Files\SL Application Cleanup"
        $64BitApps = $null
        $32BitApps = $null
        $fullQueryResults = $null
        $query = $null
        $AppsFoundToUninstall = $null
        $AppsFoundToUninstallRecheck = $null

        Write-Verbose "Setting up 32Bit powershell location to use"
        $powershellx86 = $env:SystemRoot + "\syswow64\WindowsPowerShell\v1.0\powershell.exe"

        Write-Verbose "Creating query to use for finding applications"
        $query = "select * from Win32Reg_AddRemovePrograms where (DisplayName like '%$AppToUninstall%')"
        If($KeepAppName){
            $query = $query + " and (NOT DisplayName like '$KeepAppName')"
        }

        Write-Verbose "Query to use: `"$query`""
	}
	process {
		Write-Verbose "Beginning main process loop"

        Write-Verbose "Looking for apps in 64Bit location"
        $64BitApps = get-wmiobject -Query $query | select displayname, version, ProdID | sort displayname

        Write-Verbose "Looking for apps in 32Bit location"
        $32BitApps = & $powershellx86 -args $query -command {get-wmiobject -Query $args[0] | select displayname, version, ProdID | sort displayname}

        Write-Verbose "Combining query results"
        $AppsFoundToUninstall = [Array]$64BitApps + $32BitApps | sort displayname

        Write-Verbose "Testing for apps to uninstall"
        If($AppsFoundToUninstall){
            Write-Verbose "Found instances of $AppToUninstall"
            Write-Verbose "Will now process the results and look for the app to uninstall and will uninstall"
            $AppsFoundToUninstall | select * | where {$_.Displayname -like "*$AppToUninstall*"} | select ProdID, Displayname | Foreach {
                                                    Write-Verbose "Found: $($_.Displayname)"
                                                    $installOutput = Start-Process -FilePath "C:\Windows\System32\MsiExec.exe" -ArgumentList "/x $($_.ProdID) /qn /norestart /l*v `"c:\isos\$($_.Displayname).log`"" -Wait -PassThru
                                                    #Unfortunately the wait parameter doesn't wait when used in a remote script block
                                                    #Putting in a loop to check HasExited has exited
                                                    Write-Verbose "Waiting for uninstall to complete"
                                                    do {Start-Sleep -Seconds 1 }
                                                    until ($installOutput.HasExited)
                                        }

            Write-Verbose "Resetting query result variables ready for recheck"
            $64BitApps = $null             # Stores the results of the 64Bit query
            $32BitApps = $null             # Stores the results of the 32Bit query

            Write-Verbose "Rechecking 64Bit applications"
            $64BitApps = get-wmiobject -Query $query | select displayname, version, ProdID | sort displayname

            Write-Verbose "Rechecking 32Bit applications"
            $32BitApps = & $powershellx86 -args $query -command {get-wmiobject -Query $args[0] | select displayname, version, ProdID | sort displayname}

            Write-Verbose "Combining 64Bit and 32Bit second results"
            $AppsFoundToUninstallRecheck = [Array]$64BitApps + $32BitApps | sort displayname

            If(!($AppsFoundToUninstallRecheck)){
                Write-Verbose "Applications now uninstalled"
                Write-Verbose "Creating SCCM application detection file"
                If(!(Test-Path $sccmDetectionFolder)){$temp = New-Item $sccmDetectionFolder -type directory -Force}
                $temp = New-Item "$sccmDetectionFolder\$SCCMAppName.txt" -type file -Force
                Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "===================================================================="
			    Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "`nTime of entry: $currentDate (D/M/Y) $currentTime (local)"
                Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "`nUsed for SCCM detection for application: $SCCMAppName"
                Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "`nUninstalled the following:"
                $AppsFoundToUninstall | select DisplayName | ForEach {Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "$($_.DisplayName)"}             
            }else{
                 Write-Verbose "ERROR: Applications were not uninstalled"
            }
        }Else{
            Write-Verbose "Found no instances of $AppToUninstall"
            Write-Verbose "Creating SCCM application detection file"

            If(!(Test-Path $sccmDetectionFolder)){$temp = New-Item $sccmDetectionFolder -type directory -Force}
            $temp = New-Item "$sccmDetectionFolder\$SCCMAppName.txt" -type file -Force
            Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "===================================================================="
			Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "`nTime of entry: $currentDate (D/M/Y) $currentTime (local)"
            Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "Used for SCCM detection for application: $SCCMAppName"
            Add-content "$sccmDetectionFolder\$SCCMAppName.txt" -value "`nNo applications found to uninstall"
        }
        Write-Verbose "Completed main process loop"
	}
    End {
        Write-Verbose "Function complete"
    }
}
