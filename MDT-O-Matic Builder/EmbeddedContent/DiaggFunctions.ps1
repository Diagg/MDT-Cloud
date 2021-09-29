###################################################
#
# PsFonction 5.5
#
# Compiled By Diagg/www.OSD-Couture.com
#
# My collection of usefull function   
#
#
#
# Changes : 
#			08/04/2015 v1.1 Added Copy-ItemWithProgress function
#			09/04/2015 v1.2 Added Color to log output to console
#					Added Get-CMSite function
#			14/04/2015 V1.3 Bug fix
#			08/05/2016 v1.4 Added Set-CMFolder function
#			16/06/2016 v1.5	Added Get-WindowsVersion function
#					Added Set-DeploymentEnv function
#			11/07/2016 v1.6 Bug fix	
#			06/12/2016 V1.7 Fixed ZtiUtility import when in SCCM
#			26/02/2017 V1.8 Added TestandLog-Path function
#			27/02/2017 V1.9 Added New-Shortcut function
#			12/04/2017 V2.0 Function Copy-ItemWithProgress has now more parameters
#					Copy-ItemWithProgress -source $SrcPath -dest $DestPath -arg "/Mir" 
#					-arg is not mandatory, if no arg then folder & subfolder content is also copied
#			19/04/2017 V3.0	Added a bunch of functions from Noha Swanson https://github.com/ndswanson/TaskSequenceModule
#					https://ndswanson.wordpress.com/2016/01/02/task-sequence-powershell-module/
#			19/07/2017 V3.1 Log will not be broadcasted in console during deployment
#					Finaly fixed an old bug in the logs
#			28/07/2017 V3.2 Added TestAndLog-RegistryValue
#			31/07/2017 V3.3 Added Invoke-executable by by Nickolaj Andersen
#			17/10/2017 V3.4 Added logging to BDD.log
#			27/01/2018 V3.5	Corrected a bug in the OutToConsole log function
#			31/01/2018 V3.51Logging optimized
#			16/03/2018 V3.6	Added -logpath to Int-Function so a custom log path can be set
#					Invoke-Execution rewritten to generate commande output
#			22/11/2018 V3.7 Added Set-Registry Function
#					Fixed TestAndLog-Registry
#					Fixed TestAndLog-Path
#			23/11/2018 V3.8 Added logic to load MDT module from MDT install folder
#			24/11/2018 V3.9 Reworked default log file logic
#			06/12/2018 V4.0	Variable scope is now $Script: (Not $Global)
#				        Added import-MDTSnappin function
#			16/12/2018 V4.1 Added Dump-ObjetToXml to save an object to an XML file
#					Added Dump-XmlToObject to load an object saved into XML
#			22/12/2018 V4.2	Added the PSINI module from Oliver LipKau https://github.com/lipkau/PsIni. allow easy managment of ini and inf files 
#			23/12/2018 V4.3 Added function ShowToast to display Toast notifiactions
#			26/12/2018 V4.4 Added function New-ManagedCredential and Get-ManagedCredential from BornToReboot https://github.com/BornToBeRoot/PowerShell_ManagedCredential
#			12/02/2019 V4.5 Added function Resolve-Error
#			08/03/2019 V4.6 Added Cache space to the logging function
#                       11/03/2019 V4.7 Added WMI Name space to Get-CMSite
#			18/03/2019 V4.8 Changed output of Invoke-Executable function. the return value is an array, first item of array is the return code.
#			10/04/2019 V4.9 Init-logging now support inital relocation
#					Log-ScriptEvent now support message at position 0 (no more needs to use -Value)
#			24/04/2019 V5.0 Added support for secure credencial in Invoke-executable
#			22/05/2019 V5.1	Added Function Add-XmlFragment
#			14/07/2019 V5.2 Changed Init-Function, all preset variables are now enumerated
#			01/08/2019 V5.3 Added Function Get-ComObject
#			11/07/2019 V5.4 Fixed various Bugs
#			14/08/2021 V5.5 made Quiet mode really quiet in function TestAndLog-RegistryValue
#
#           
###################################################

Function Init-Logging
	{
		# Create the default log file.
		# This function must be executed once to avoid error and to make logs working.
		# If no full Log file path is specified, a default log file is created in the script folder
		# The script will also create his folder path if it does not exists
		# ex: $Script:LogFile = Init-Logging -logpath $("C:\ProgramData\MDT-Manager\$CurrentUser\Logs" + "\" + $CurrentScriptName.replace("ps1","log"))
	
	    Param (
	            [parameter(Mandatory = $false,Position=0)]
				[Alias('logPath')]
	            [string]$oPath = ($CurrentScriptPath + "\" + $CurrentScriptName.replace("ps1","log"))
			)
			
		$oFolderPath = Split-Path $oPath
		If (-not (test-path $oFolderPath))
		    {
		        New-Item -Path $oFolderPath -ItemType Directory -Force|out-null
		        If (-not (test-path $oFolderPath)){Write-Error "[ERROR] Unable to create $oFolderPath, Aborting!!!!"; Exit}
		    }
			
		Log-ScriptEvent -NewLog $oPath -value "***************************************************************************************************"
		Log-ScriptEvent -NewLog $oPath -value "Started processing time: [$([DateTime]::Now)]"
		Log-ScriptEvent -NewLog $oPath -value "Script Name: $CurrentScriptName"
		Log-ScriptEvent -NewLog $oPath -value "Log location: $oPath"
		Log-ScriptEvent -NewLog $oPath -value "***************************************************************************************************"
		Log-ScriptEvent -NewLog $oPath -value " "
		Log-ScriptEvent -NewLog $oPath -value "***************************************************************************************************"
		Return $oPath
	}


Function Relocate-Logging
	{
        # Move logging to another location.
        # The path will be created if it does not exists
	# Ex:  $Script:LogFile = Relocate-Logging -path $NewLogPath
		
		Param (
            [parameter(Mandatory=$true,Position=0)]
			[string]$Path
			)

       
        $testPath = TestAndLog-Path $Path
        If ($testPath -eq $false)
            {
                New-Item -Path $Path -ItemType Directory -force
                $testPath = TestAndLog-Path $Path
                If ($testPath -eq $false)
					{
						Log-ScriptEvent -value "[ERROR] Unable to move log file to $Path aborting!!!" -Severity 2
						Return $LogFile
					}
            }

        
        If ($testPath -eq $true)
            {
                Log-ScriptEvent "***************************************************************************************************"
				Log-ScriptEvent "Old log path was [$LogFile]"
                Move-Item -path $LogFile -Destination $Path -Force
                $NewFullLogPath =  ($Path + "\" + $CurrentScriptName.replace("ps1","log"))
                If (test-Path $NewFullLogPath)
					{	
						Log-ScriptEvent -NewLog $NewFullLogPath -Value "Relocated log path is [$NewFullLogPath]"
						Log-ScriptEvent -NewLog $NewFullLogPath -Value "***************************************************************************************************"
						Return $NewFullLogPath
					} 
            }
    }

Function Get-DiaggShortName
    {
	    # This function allow powershell to manage files or folders with special char-set like "[]"
	    # Borrowed to http://stackoverflow.com/questions/16995359/get-childitem-equivalent-of-dir-x
	    # input parameter = file ou folder
	    # Return path in 8.3 format !
	    # see link for recusive exemple.

        Param (
	            [parameter(Mandatory = $true,Position=0)]
	            [string]$oPath
			)

        $fso = New-Object -ComObject Scripting.FileSystemObject

        if (Test-Path -literalpath $oPath)
            {
                if((Get-Item -literalpath $opath).psiscontainer) 
                    {Return $fso.GetFolder($opath).ShortPath}
                else 
                    {Return $fso.GetFile($opath).ShortPath} 
            }
    }
	
	
Function Log-ScriptEvent 
	{

		##########################################################################################################
		<#

		This Function by Ian Farr : https://gallery.technet.microsoft.com/scriptcenter/Log-ScriptEvent-Function-ea238b85

		.SYNOPSIS
		   Log to a file in a format that can be read by Trace32.exe / CMTrace.exe 

		.DESCRIPTION
		   Write a line of data to a script log file in a format that can be parsed by Trace32.exe / CMTrace.exe

		   The severity of the logged line can be set as:

		        1 - Information
		        2 - Warning
		        3 - Error

		   Warnings will be highlighted in yellow. Errors are highlighted in red.

		   The tools to view the log:

		   SMS Trace - http://www.microsoft.com/en-us/download/details.aspx?id=18153
		   CM Trace - Installation directory on Configuration Manager 2012 Site Server - <Install Directory>\tools\

		.EXAMPLE
		   Log-ScriptEvent c:\output\update.log "Application of MS15-031 failed" Apply_Patch 3

		   This will write a line to the update.log file in c:\output stating that "Application of MS15-031 failed".
		   The source component will be Apply_Patch and the line will be highlighted in red as it is an error 
		   (severity - 3).

		#>
		##########################################################################################################



		#Define and validate parameters
		[CmdletBinding()]
		Param(
		      #Path to the log file
		      [parameter(Mandatory=$False)]
		      [String]$NewLog = $LogFile,

		      #The information to log
		      [parameter(Mandatory=$True,Position=0)]
		      [String]$Value,

		      #The source of the error
		      [parameter(Mandatory=$False)]
		      [String]$Component = $CurrentScriptName,

		      #The severity (1 - Information, 2- Warning, 3 - Error)
		      [parameter(Mandatory=$False)]
		      [ValidateRange(1,3)]
		      [Single]$Severity = 1,
			  
			  #Also output to console ($True or $False)
		      [parameter(Mandatory=$False)]
		      [bool]$OutToConsole = $True
			  			  
		      )

		#Obtain UTC offset
		$DateTime = New-Object -ComObject WbemScripting.SWbemDateTime 
		$DateTime.SetVarDate($(Get-Date))
		$UtcValue = $DateTime.Value
		$UtcOffset = $UtcValue.Substring(21, $UtcValue.Length - 21)


		#Create the line to be logged
		$LogLine =  "<![LOG[$Value]LOG]!>" +`
		            "<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
		            "date=`"$(Get-Date -Format M-d-yyyy)`" " +`
		            "component=`"$Component`" " +` 
		            "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
		            "type=`"$Severity`" " +`
		            "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
		            "file=`"`">"

		#Write the line to the passed log file
        $oLogLine = $LogLine
        :PushContent Do
            {
		        Try
			        {
                        If ( -not [String]::IsNullOrWhiteSpace($Global:LogTemp))
                            {
                               $oLogLine = $Global:LogTemp + [Environment]::NewLine + $oLogLine
                               $Global:LogTemp = [String]::Empty
                            }
                        
                        Add-Content -Path $NewLog -Value $oLogLine -ErrorAction Stop
                        $ContentFlushed = $True
                    } 
		        Catch
			        {
				        $ContentFlushed = $False
                        While ((TestAndLog-Path -Path $NewLog -NoLogging) -eq $False)
					        { 
						        Start-Sleep -Milliseconds 200
						        $oWait= $oWait +200
						        If ($owait -ge 2000)
                                    {
                                        $Global:LogTemp = $oLogLine
                                        Break PushContent
                                    }
					        }
                    }
            }Until($ContentFlushed -eq $True)
		
		
        #Write the Line to BDD.log
		If ($OSD_Env.TSenv -eq $True -and  (!([string]::IsNullOrEmpty($OSD_Env.BDDLog))))
			{
                $oLogLine = $LogLine
                :PushBDDContent Do
                    {
		                Try
			                {
                                If ( -not [String]::IsNullOrWhiteSpace($Global:LogTemp))
                                    {
                                       $oLogLine = $Global:LogTemp + [Environment]::NewLine + $oLogLine
                                       $Global:LogTemp = [String]::Empty
                                    }
                                Add-Content -Path $($OSD_Env.BDDLog) -Value $oLogLine -ErrorAction Stop
                                $ContentFlushed = $True
                            } 
		                Catch
			                {
				                $ContentFlushed = $False
                                While ((TestAndLog-Path -Path $($OSD_Env.BDDLog) -NoLogging) -eq $False)
					                { 
						                Start-Sleep -Milliseconds 200
						                $oWait= $oWait +200
						                If ($owait -ge 2000)
                                            {
                                                $Global:LogTemp = $oLogLine
                                                Break PushBDDContent
                                            }
					                }
				
			                }
                    }Until($ContentFlushed -eq $True)
			}
		
		

		If ($OutToConsole)
			{
				If (($script:OSD_Env.TSenv -eq $True -and $script:OSD_Env.IsStandAlone -eq $true) -or $script:OSD_Env.TSenv -eq $False)
					{
						switch ($Severity)
							{
								1 {Write-Host ((Get-Date -Format HH:mm:ss)+ " - " + $Value); break}
								2 {Write-Host ((Get-Date -Format HH:mm:ss)+ " - " + $Value) -ForegroundColor black -BackgroundColor Yellow; break}
								3 {Write-Host ((Get-Date -Format HH:mm:ss)+ " - " + $Value) -ForegroundColor Red; break}
							}
					}		
			}
	}	
	

Function Copy-ItemWithProgress
	{
		<#
		.SYNOPSIS
		RoboCopy with PowerShell progress.

		.DESCRIPTION
		Performs file copy with RoboCopy. Output from RoboCopy is captured,
		parsed, and returned as Powershell native status and progress.

		.PARAMETER RobocopyArgs
		List of arguments passed directly to Robocopy.
		Must not conflict with defaults: /ndl /TEE /Bytes /NC /nfl /Log

		.OUTPUTS
		Returns an object with the status of final copy.
		REMINDER: Any error level below 8 can be considered a success by RoboCopy.

		.EXAMPLE
		C:\PS> .\Copy-ItemWithProgress c:\Src d:\Dest

		Copy the contents of the c:\Src directory to a directory d:\Dest
		With the /e switch, files, folder and subfolders from the root of c:\src are copied.

		.EXAMPLE
		C:\PS> .\Copy-ItemWithProgress '"c:\Src Files"' d:\Dest /mir /xf *.log -Verbose

		Copy the contents of the 'c:\Name with Space' directory to a directory d:\Dest
		/mir and /XF parameters are passed to robocopy, and script is run verbose

		.LINK
		http://keithga.wordpress.com/2014/06/23/copy-itemwithprogress

		.NOTES
		By Keith S. Garner (KeithGa@KeithGa.com) - 6/23/2014
		With inspiration by Trevor Sullivan @pcgeek86

		#>

		[CmdletBinding()]
		param(
			[Parameter(Mandatory = $true)] 
			[string[]] $Source,
			[Parameter(Mandatory = $true)] 
			[string[]] $Dest,			
			[Parameter(Mandatory = $false,ValueFromRemainingArguments=$true)] 
			[string[]] $Args="/E"
		)
		
		#remove last char if it's "\"
		If ($Source.EndsWith("\")) {$source = $source.Substring(0,$source.Length-1)}
		If ($Dest.EndsWith("\")) {$Dest = $Dest.Substring(0,$Dest.Length-1)}
		
		# Add quote to paths
		$Source = [char]34 + $Source + [char]34
		$Dest = [char]34 + $Dest + [char]34
		
		#Rebuild arg
		$RobocopyArgs = $Source + " " + $Dest + " " + $Args + " " 
		
		$ScanLog  = [IO.Path]::GetTempFileName()
		$RoboLog  = [IO.Path]::GetTempFileName()
		$ScanArgs = $RobocopyArgs + "/ndl /TEE /bytes /Log:$ScanLog /nfl /L".Split(" ")
		$RoboArgs = $RobocopyArgs + "/ndl /TEE /bytes /Log:$RoboLog /NC".Split(" ")

		# Launch Robocopy Processes
		Log-ScriptEvent -value ("Robocopy Scan:`n " + ($ScanArgs -join " ")) -Component "Function-Library(Copy-ItemWithProgress)" -Severity  1 -OutToConsole $True
		Log-ScriptEvent -value ("Robocopy Full:`n " + ($RoboArgs -join " ")) -Component "Function-Library(Copy-ItemWithProgress)" -Severity  1 -OutToConsole $True
		Log-ScriptEvent -value "Log file Scan : $ScanLog" -Component "Function-Library(Copy-ItemWithProgress)" -Severity  1 -OutToConsole $True
		Log-ScriptEvent -value "Log file Copy : $RoboLog" -Component "Function-Library(Copy-ItemWithProgress)" -Severity  1 -OutToConsole $True
		$ScanRun = start-process robocopy -PassThru -WindowStyle Hidden -ArgumentList $ScanArgs
		$RoboRun = start-process robocopy -PassThru -WindowStyle Hidden -ArgumentList $RoboArgs

		# Parse Robocopy "Scan" pass
		$ScanRun.WaitForExit()
		$LogData = get-content $ScanLog
		if ($ScanRun.ExitCode -ge 10)
			{
				$LogData|out-string|Write-Error
				throw "Robocopy $($ScanRun.ExitCode)"
			}
		$FileSize = [regex]::Match($LogData[-4],".+:\s+(\d+)\s+(\d+)").Groups[2].Value
		Log-ScriptEvent -value ("Robocopy Bytes: $FileSize `n" +($LogData -join "`n")) -Component "Function-Library(Copy-ItemWithProgress)" -Severity  1 -OutToConsole $True

		# Monitor Full RoboCopy
		while (!$RoboRun.HasExited)
		{
			$LogData = get-content $RoboLog
			$Files = $LogData -match "^\s*(\d+)\s+(\S+)"
		    if ($Files -ne $Null )
			    {
				    $copied = ($Files[0..($Files.Length-2)] | %{$_.Split("`t")[-2]} | Measure -sum).Sum
				    if ($LogData[-1] -match "(100|\d?\d\.\d)\%")
					    {
						    write-progress Copy -ParentID $RoboRun.ID -percentComplete $LogData[-1].Trim("% `t") $LogData[-1]
						    $Copied += $Files[-1].Split("`t")[-2] /100 * ($LogData[-1].Trim("% `t"))
					    }
				    else
					    {
						    write-progress Copy -ParentID $RoboRun.ID -Completed
					    }
					$PercentComplete = [math]::min(100,(100*$Copied/[math]::max($Copied,$FileSize)))
					write-progress ROBOCOPY -ID $RoboRun.ID -PercentComplete $PercentComplete $Files[-1].Split("`t")[-1] 
			    }
		}

		write-progress Copy -ParentID $RoboRun.ID -Completed
		write-progress Copy -ID $RoboRun.ID -Completed

		# Parse full RoboCopy pass results, and cleanup
		(get-content $RoboLog)[-50..-2] | out-string | Write-Verbose
		[PSCustomObject]@{ ExitCode = $RoboRun.ExitCode }
		remove-item $RoboLog, $ScanLog
	}


Function Get-CMSite
	{

		# this funtion will return an object with properties like SCCM Site Name and Sccm server site Name. It will also import the SCCM powershell module.
		# thanks and respect to Andrew Barns for great inspiration

	    #Load the ConfigurationManager Module
	    If (test-path "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1" )
	        {
	            If (!(Get-Module ConfigurationManager))
	                {
	                    Log-ScriptEvent -value "Importing SCCM Module." -Component "Function-Library(Get-CMSite)" -Severity  1 -OutToConsole $True
						Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
	                }
	        }
	    Else
	        {
	            Log-ScriptEvent -value "No Powershell module found !!!" -Component "Function-Library(Get-CMSite)" -Severity  3 -OutToConsole $True
				Log-ScriptEvent -value "SCCM Console not installed or script launched with insuffisant access right. Exiting !!!" -Component "Function-Library" -Severity  3 -OutToConsole $True
	            Exit
	        }

	    # Check if the SCCM drive Exist 
	    $CCMDrive = (Get-PSDrive -PsProvider CMSITE).Name
	    If (!([string]::IsNullOrEmpty($CCMDrive)))
	        {
	            Log-ScriptEvent -value ("SCCM Drive Found with name : " + $CCMDrive) -Component "Function-Library" -Severity  1 -OutToConsole $True
				[PSCustomObject]@{ SiteCode = $CCMDrive ; SiteDrive = ($CCMDrive + ":") ; SiteServer = (get-psdrive $CCMDrive).root ; WMInameSpace = ("ROOT\SMS\Site_" + $CCMDrive) }
	        }
	    Else
	        {
	            Log-ScriptEvent -value "Unable to find SCCM Drive. Exiting !!!" -Component "Function-Library" -Severity  3 -OutToConsole $True
	            Exit
	        }

	}
	

Function Set-CMFolder
	{
	
		#This function will create folders into SCCM Console
		# Usage Set-CMFolder -Path <SCCM console Path without site name>
		# Exemple Set-CMFolder -Path "\Package\12 - Deploiement OS\NEDUGO"
	
		#Define and validate parameters
		[CmdletBinding()]
		Param(
			      #Path to the SCCM Folder
			      [parameter(Mandatory=$True,Position=0)]
			      [String]$Path

		    	)
	
	
		# Save Current location
		$CurrentLocation = Get-Location
				
		# Relocate to SCCM Drive
		CD $CMSiteInfo.SiteDrive
		
		# Rebuild fill patch
		$Path = $CMSiteInfo.SiteDrive + $Path

		# create folders in SCCM console Driver's folder
		Log-ScriptEvent -value "Creating folders in SCCM Console with this path $Path "  -Component $CurrentScriptName -Severity  1 -OutToConsole $True
		$FinalFolder =""
		$subFolders = $Path.Split("\")
		ForEach ($Folder in $subFolders)
			{
				$FinalFolder = $FinalFolder + $Folder + "\"
				
				If (!(Test-Path $FinalFolder.substring(0,$FinalFolder.Length - 1)))
				    {
						Log-ScriptEvent -value "Creating new sccm folder : $FinalFolder" -Component $CurrentScriptName -Severity  1 -OutToConsole $True
						New-item -path $FinalFolder

				    }
			}
			
		# relocate back to previous location
		CD $CurrentLocation
	}	
	
	
Function Get-WindowsVersion
	{
		# this funtion will return an object with properties like :
		# full Windows version number as string (.fullNum) ex: 10.0.10586
		# Short Windows version number as a number (.MiniNum) ex: 10
		# SKU Name (.SKU) ex: Windows 10
		# Short SKU Name (.MinSKU) ex: Win10
		# Os Architecture (.Arch) ex: x64
		# Build Number as a number(Build) ex: 10586
		# Edition of Windows (.edition) ex: Microsoft Windows 10 Enterprise
		# Service Pack Level (.ServicePack) ex: SP1
		# Servicing CBB (.IsCBB) ex: $True
		# Servicing LSTB (.IsLTSB) ex: $False
		
		$SysWMI = Get-WmiObject win32_operatingSystem
        
        $FullNum =  $SysWMI.Version
        $MiniNum= $FullNum.Split(".")[0] + "." + $FullNum.Split(".")[1] 
        switch ($MiniNum)
            {
             "10.0" { [int]$MiniNum = 10 ; Break}
             "6.1" { [int]$MiniNum = 7 ; Break}
             "6.2" { [int]$MiniNum = 8 ; Break}
             "6.3" { [single]$MiniNum = 8,1 ; Break}
            }

		$SKU = ("Windows " + ($FullNum.Split(".")[0]))
		$MiniSKU = ("Win" + ($FullNum.Split(".")[0]))
		[int]$Build = [Convert]::ToInt32($SysWMI.BuildNumber,10)
		
		$Arch = $SysWMI.OSArchitecture
		If ($Arch -like "*64*") {$Arch = "x64"} Else {$Arch = "x86"}
		
        $Edition = TestAndLog-RegistryValue -path 'Hklm:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\EditionID' -Action quiet
		
        $Friendlyversion = TestAndLog-RegistryValue -path 'Hklm:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DisplayVersion' -Action quiet
        If ($Friendlyversion -eq $false){$Friendlyversion = (TestAndLog-RegistryValue -path 'Hklm:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ReleaseID' -Action quiet)}

        $ProductName = TestAndLog-RegistryValue -path 'Hklm:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName' -Action quiet
		
		If ($MiniNum -ge 10)
			{
				If (($ProductName).contains('LTSB'))
				   	{$IsLTSB = $true; $IsCBB = $False} 
                Else 
                    {$IsLTSB = $False; $IsCBB = $True}
			}
		Else
			{$IsLTSB = $False; $IsCBB = $False}
				
		[PSCustomObject]@{ FullNum = $FullNum ; MiniNum = $MiniNum ; SKU = $SKU ; MiniSKU = $MiniSKU ; Arch = $Arch ; Build = $Build ; Edition = $Edition ; Friendlyversion = $Friendlyversion ; ProductName =$ProductName ; IsLTSB = $IsLTSB ; IsCBB = $IsCBB }
		
	}

	
Function Set-DeploymentEnv
	{
		# Warning: Log-ScriptEvent is not yet initialized and can't be used at this stage !!!! 
	
		# We just tried to import the module, if it's still not there, try to find it on the system disk
        If (!((Get-Module).name -eq "ZTIUtility"))	
			{
                If (Test-path 'C:\Program Files\Microsoft Deployment Toolkit\Templates\Distribution\Tools\Modules\ZTIUtility\ZTIUtility.psm1')
                    {
                        Import-Module 'C:\Program Files\Microsoft Deployment Toolkit\Templates\Distribution\Tools\Modules\ZTIUtility\ZTIUtility.psm1' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                        $IsStandAlone = $true
                        $IsMDT = $false
                    }
                ElseIf (Test-path 'C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\ZTIutility\ZTIUtility.psm1')
                    {
                        Import-Module 'C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\ZTIutility\ZTIUtility.psm1' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                        $IsStandAlone = $true
                        $IsMDT = $false
                    }
				Else
					{
						# Loop thru each drives, to detect USB media or Network Path
						ForEach ($Item in 97..(97+25))
							{
								$Drive = [char]$Item
								If (Test-Path $($Drive + ":\Deploy\Tools\Modules\ZTIUtility\ZTIUtility.psm1"))
									{
                                        Import-Module $($Drive + ":\Deploy\Tools\Modules\ZTIUtility\ZTIUtility.psm1") -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                        $IsStandAlone = $false
                                        $IsMDT = $True
                                        Break
                                     }
								ElseIf (Test-Path $($Drive + ":\Tools\Modules\ZTIUtility\ZTIUtility.psm1"))
									{
                                        Import-Module $($Drive + ":\Tools\Modules\ZTIUtility\ZTIUtility.psm1") -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                                        $IsStandAlone = $false
                                        $IsMDT = $True
                                        Break
                                    }	
							}
					}
            }


        #Set TSENV
        If (!((Get-Module).name -eq "ZTIUtility"))
            {
                $TSenv = $False
				$IsStandAlone = $false
			    $IsMDT = $false
            }
        Else
            {
				##== Verfiy that the drive is loaded and not empty
				If((Get-PSDrive).name -eq "TSenv"){If ((Get-childitem TSenv:).count -gt 10) {$TSenv = $True} Else {$TSenv = $False}} 

                ##== Check if we are running within MDT or SCCM
                If((@("*MININT\Tools*","*X:\Deploy\Tools\X64\TsmBootstrap.exe*").Where{(Get-Process -Name "TsmBootstrap" -ErrorAction SilentlyContinue).path -like $_}) -and $TSenv -eq $True) 
                    {
                        $IsMDT = $True
                        $IsSCCM = $false
                        $IsStandAlone = $false
                    }
                ElseIf($TSenv -eq $True ) 
                    {
                        $IsMDT = $False
                        $IsSCCM = $True
                        $IsStandAlone = $false
                    }
                Else
                    {
                        $IsMDT = $false
                        $IsSCCM = $false
                        $IsStandAlone = $true
                    }
            }

        ##== Add Module to PS path
        if (($env:PSModulePath -split ";") -notcontains $((Get-Module ZTIUtility).ModuleBase)) {$env:PSModulePath += ";$((Get-Module ZTIUtility).ModuleBase)"}


        ##== set additionnal Infos
		If ($TSenv -eq $True) 
			{
                #Get Task Sequence Name
				[xml]$TSXml = $tsenv:_SMSTSTaskSequence
				$TSName = $TSXml.sequence.Name
								
				#Get BDD.Log for MDT or Stand Alone
				If ($IsMDT -eq $true -or $IsStandAlone -eq $True)
					{
						$oDrives = (get-psdrive| where Provider -like "*FileSystem*" ).Root 
						ForEach ( $oDrv in $oDrives )
							{
								$oBddLog = ($oDrv + "MININT\SMSOSD\OSDLOGS\BDD.log")
								If (test-path $oBddLog ){break}
							}
					}


				#Get smsts.Log for SCCM
				If ($IsSCCM -eq $true)
					{
						$oDrives = (get-psdrive| where Provider -like "*FileSystem*" ).Root 
						ForEach ( $oDrv in $oDrives )
							{
								#If SCCM Clienet is not yet installed
                                $oBddLog = ($oDrv + "_SMSTaskSequence\Logs\Smstslog\smsts.log")
								If (test-path $oBddLog ){break}

								#If SCCM client is Installed
                                $oBddLog = ($oDrv + "windows\ccm\logs\Smstslog\smsts.log")
								If (test-path $oBddLog ){break}
							}
					}
            }
        Else
            {
                $oBddLog = $false
                $TSName = $false
                $TSXml = $false
            }


        # Check if WinPE
        $IsWinPE = (Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlset\Control\MiniNT)

        # Check If UEFI
        $IsUEFI = ($env:firmware_type -eq "UEFI")


        #Create Returning Object
        [PSCustomObject]@{ TaskSequenceName = $TSName  ; IsSCCM = $IsSCCM ; IsMDT = $IsMDT ; IsStandAlone = $IsStandAlone ; TSenv = $TSenv ; TaskSequenceXML = $TSXml ; BDDLog = $oBddLog ; IsWinPE = $IsWinPE ; IsUEFI = $IsUEFI }
	}
	

Function Set-SCCMDeveloppementEnv
	{
	
		# Warning: Log-ScriptEvent is not yet initialized and can't be used at this stage !!!! 
		# this function will check if We are working on an 'SCCM capable machine' (a machine where SCCM console is at least installed
		# If yes, will load the ConfigurationManager Module
		# and will retrune the state of the enviroment and the state of the module
		
		$software = "System Center Configuration Manager Console"
		$installed = (Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where { $_.DisplayName -like "*$software*" })

		If($installed) 
			{
				If ((Get-Module ConfigurationManager).name -eq $null) 
					{
						Import-Module (Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -Force -ErrorAction SilentlyContinue
						If ((Get-Module ConfigurationManager).name -eq $null)
							{
								Write-Host "Ps Module Registration failed at first attempt, trying a second time !!"
								Import-Module (Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -Force -ErrorAction SilentlyContinue
								If ((Get-Module ConfigurationManager).name -eq $null)
									{
										Write-Host "[ERROR] Unable to load ConfigManager Powershell Module, aborting!!!"
										[PSCustomObject]@{ SCCMDevEnv = $true  ; SCCMDevModule = $false ; SCCMDevModVer = 0 ; SCCMDevModPath = $false}
										Return
									}
								Else
									{
										$ModuleInfo = Get-Module ConfigurationManager
										If (-not ([String]::IsNullOrWhiteSpace($ModuleInfo)))
											{
												Write-Host "ConfigManager Powershell Module loaded succefully !!"
												[PSCustomObject]@{ SCCMDevEnv = $true  ; SCCMDevModule = $true ; SCCMDevModVer = $ModuleInfo.Version ; SCCMDevModPath = $ModuleInfo.Path} 
												Return
											}
										Else
											{
												Write-Host "[ERROR] ConfigManager Powershell Module loaded, but does not works properly, aborting!!!"
												[PSCustomObject]@{ SCCMDevEnv = $true  ; SCCMDevModule = $false ; SCCMDevModVer = 0 ; SCCMDevModPath = $false}
												Return
											}									
									}
							}
					}
			}
		Else
			{[PSCustomObject]@{ SCCMDevEnv = $false  ; SCCMDevModule = $false ; SCCMDevModVer = 0  ; SCCMDevModPath = $false} ; Return}
	}	
	
	
	
	
Function TestAndLog-Path
	{
	
		#Define and validate parameters
		[CmdletBinding()]
		Param(
			      #Path to verify
			      [parameter(Mandatory=$True)]
			      [String]$Path,
				  
				  #Action to log
				  [ValidateSet("created","modified","checked","moved","copied")]
				  [String]$Action="checked",

				  #Log stuffs or Not
				  [Switch]$NoLogging


		    	)
		
		If (Test-Path $Path)
			{
				
                #Verify that this is a file and not a directory
                If ([System.IO.File]::Exists($Path)) 
                    {
                        # Check if path is locked		
				        $oFile = New-Object System.IO.FileInfo $Path

				        try 
					        {
		    			        #Start-Sleep -Seconds 2
                                $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
		    			        if ($oStream) {$oStream.Close() ;$oStream.Dispose()}
                                If ($NoLogging -eq $false){Log-ScriptEvent -value "File at $Path was $Action Sucessfully" -Severity 1}
						        return $True
		  			        } 
				        catch 
					        {
		    			        Get-ChildItem $path -ErrorAction SilentlyContinue|out-null
                                if ($Error[0].Exception -is [System.UnauthorizedAccessException])
                                    {

                                        If ($NoLogging -eq $false){Log-ScriptEvent -value "File at $Path has access denied !" -Severity 2}
						                return $false
                                    }
                                Else
                                    {
                                        # file is locked by a process.
						                If ($NoLogging -eq $false){Log-ScriptEvent -value "File at $Path is locked and can't be processed !" -Severity 2}
						                return $false
                                    }
		  			        }				
				    }
                
                If ($NoLogging -eq $false){Log-ScriptEvent -value "Folder at $Path was $Action Sucessfully" -Severity 1}
				return $True
			}
		Else	
			{
				If ($Logging -eq $false){Log-ScriptEvent -value "Error, $Path not found !!!" -Severity  2}
				Return $false
			}
	}
	
	
function TestAndLog-RegistryValue 
	{

		param (

				#Path to Verify
				[parameter(Mandatory=$true)]
				[ValidateNotNullOrEmpty()]$Path,

				#Registry Value to verify
				[parameter(Mandatory=$false)]
				[String]$Value,
				 
				#Action to log
				[ValidateSet("created","modified","checked","quiet")]
				[String]$Action="checked"
			 
			)

        $CheckPath = Test-path $Path

        # This is a hint that we should investigate a bit further
        If([string]::IsNullOrWhiteSpace($Value) -and $CheckPath -eq $false)
            {
                
                $keyPath = split-path $path 
                $KeyProperty = split-path $path -Leaf

                $CheckPath = Test-path $keyPath
                If($CheckPath -eq $false)
                    {
						If (-not($Action -eq "quiet")){Log-ScriptEvent "[WARNING] Unable to check $Path, The path does not exist" -Severity  2}
						Return $false
                    } 
                        
                $KeyValue = (Get-ItemProperty -Path $keyPath -Name $KeyProperty -ErrorAction SilentlyContinue).$KeyProperty

                If([string]::IsNullOrWhiteSpace($KeyValue))
                    {
                        If (-not($Action -eq "quiet")){Log-ScriptEvent "[WARNING] Unable to check $Path, The path returned no value" -Severity  2}
						Return $false
                    }
                Else
                    {
                        If (-not($Action -eq "quiet")){Log-ScriptEvent "Reg Value $KeyValue from $Path was $Action Sucessfully"}
                        Return $KeyValue
                    }
            }
        Else
            {
                If (-not($Action -eq "quiet")){Log-ScriptEvent "[WARNING] Unable to check $Path, The path returned no value" -Severity  2}
				Return $false
            }
	
		
		If ($CheckPath -and (-not([string]::IsNullOrWhiteSpace($Value))) )
			{
                $KeyValue = (Get-ItemProperty -Path $Path -Name $value -ErrorAction SilentlyContinue).$value

                If([string]::IsNullOrWhiteSpace($KeyValue))
                    {
                        If (-not($Action -eq "quiet")){Log-ScriptEvent "[WARNING] Unable to check $Path\$Value, The path returned no value" -Severity  2}
					    Return $false
                    }
                Else
                    {
                        If (-not($Action -eq "quiet")){Log-ScriptEvent "Reg Value $KeyValue from $Path\$value was $Action Sucessfully"}
                        Return $KeyValue
                    }
    		}


		If ($CheckPath -and [string]::IsNullOrWhiteSpace($Value) )
			{
                If (-not($Action -eq "quiet")){Log-ScriptEvent "Reg Path from $Path was $Action Sucessfully"}
                Return $True
       		}
	}	
	
	

Function New-Shortcut 
	{ 
		<#   
		.SYNOPSIS   
		    This script is used to create a  shortcut.         
		.DESCRIPTION   
		    This script uses a Com Object to create a shortcut. 
		.PARAMETER Path 
		    The path to the shortcut file.  .lnk will be appended if not specified.  If the folder name doesn't exist, it will be created. 
		.PARAMETER TargetPath 
		    Full path of the target executable or file. 
		.PARAMETER Arguments 
		    Arguments for the executable or file. 
		.PARAMETER Description 
		    Description of the shortcut. 
		.PARAMETER HotKey 
		    Hotkey combination for the shortcut.  Valid values are SHIFT+F7, ALT+CTRL+9, etc.  An invalid entry will cause the  
		    function to fail. 
		.PARAMETER WorkDir 
		    Working directory of the application.  An invalid directory can be specified, but invoking the application from the  
		    shortcut could fail. 
		.PARAMETER WindowStyle 
		    Windows style of the application, Normal (1), Maximized (3), or Minimized (7).  Invalid entries will result in Normal 
		    behavior. 
		.PARAMETER Icon 
		    Full path of the icon file.  Executables, DLLs, etc with multiple icons need the number of the icon to be specified,  
		    otherwise the first icon will be used, i.e.:  c:\windows\system32\shell32.dll,99 
		.PARAMETER admin 
		    Used to create a shortcut that prompts for admin credentials when invoked, equivalent to specifying runas. 
		.NOTES   
		    Author        : Rhys Edwards 
		    Email        : powershell@nolimit.to   
		.INPUTS 
		    Strings and Integer 
		.OUTPUTS 
		    True or False, and a shortcut 
		.LINK   
		    Script posted over:  N/A   
		.EXAMPLE   
		    New-Shortcut -Path c:\temp\notepad.lnk -TargetPath c:\windows\notepad.exe     
		    Creates a simple shortcut to Notepad at c:\temp\notepad.lnk 
		.EXAMPLE 
		    New-Shortcut "$($env:Public)\Desktop\Notepad" c:\windows\notepad.exe -WindowStyle 3 -admin 
		    Creates a shortcut named Notepad.lnk on the Public desktop to notepad.exe that launches maximized after prompting for  
		    admin credentials. 
		.EXAMPLE 
		    New-Shortcut "$($env:USERPROFILE)\Desktop\Notepad.lnk" c:\windows\notepad.exe -icon "c:\windows\system32\shell32.dll,99" 
		    Creates a shortcut named Notepad.lnk on the user's desktop to notepad.exe that has a pointy finger icon (on Windows 7). 
		.EXAMPLE 
		    New-Shortcut "$($env:USERPROFILE)\Desktop\Notepad.lnk" c:\windows\notepad.exe C:\instructions.txt 
		    Creates a shortcut named Notepad.lnk on the user's desktop to notepad.exe that opens C:\instructions.txt  
		.EXAMPLE 
		    New-Shortcut "$($env:USERPROFILE)\Desktop\ADUC" %SystemRoot%\system32\dsa.msc -admin  
		    Creates a shortcut named ADUC.lnk on the user's desktop to Active Directory Users and Computers that launches after  
		    prompting for admin credentials 
		#> 
		 
		[CmdletBinding()] 
		param( 
		    [Parameter(Mandatory=$True,  ValueFromPipelineByPropertyName=$True,Position=0)]  
		    [Alias("File","Shortcut")]  
		    [string]$Path, 
		 
		    [Parameter(Mandatory=$True,  ValueFromPipelineByPropertyName=$True,Position=1)]  
		    [Alias("Target")]  
		    [string]$TargetPath, 
		 
		    [Parameter(ValueFromPipelineByPropertyName=$True,Position=2)]  
		    [Alias("Args","Argument")]  
		    [string]$Arguments, 
		 
		    [Parameter(ValueFromPipelineByPropertyName=$True,Position=3)]   
		    [Alias("Desc")] 
		    [string]$Description, 
		 
		    [Parameter(ValueFromPipelineByPropertyName=$True,Position=4)]   
		    [string]$HotKey, 
		 
		    [Parameter(ValueFromPipelineByPropertyName=$True,Position=5)]   
		    [Alias("WorkingDirectory","WorkingDir")] 
		    [string]$WorkDir, 
		 
			[Parameter(ValueFromPipelineByPropertyName=$True,Position=6)]
			[ValidateSet("Normal","Maximized","Minimized")]
			[String]$WindowStyle="Normal", 
		 
		    [Parameter(ValueFromPipelineByPropertyName=$True,Position=7)]   
		    [string]$Icon, 
		 
		    [Parameter(ValueFromPipelineByPropertyName=$True)]   
		    [switch]$admin 
		) 
		 
		 
		Process 
			{ 
		 
		  		If (!($Path -match "^.*(\.lnk)$")){$Path = "$Path`.lnk"} 
		  		[System.IO.FileInfo]$Path = $Path 
		
			  	Try 
					{ 
			    		If (!(Test-Path $Path.DirectoryName)) 
							{md $Path.DirectoryName -ErrorAction Stop | Out-Null} 
			  		} 
				Catch 
					{ 
			    		Log-ScriptEvent -value "Unable to create $($Path.DirectoryName), shortcut cannot be created"  -Component $CurrentScriptName -Severity 2 -OutToConsole $True
			    		Return $false 
			    		Break 
			  		} 
		 
				# Convert Window size to integer
				if( $WindowStyle -like "Normal" ) { [Int]$WindowStyle = 1 }
				if( $WindowStyle -like "Maximized" ) { [Int]$WindowStyle = 3 }
				if( $WindowStyle -like "Minimized" ) { [Int]$WindowStyle = 7 }

		 
				# Define Shortcut Properties 
				$WshShell = New-Object -ComObject WScript.Shell 
				$Shortcut = $WshShell.CreateShortcut($Path.FullName)
				if($Hotkey.Length -gt 0 ) { $Shortcut.HotKey = $Hotkey }
				if($Arguments.Length -gt 0 ) { $Shortcut.Arguments = $Arguments }
				if($Description.Length -gt 0 ) { $Shortcut.Description = $Description }
				if($WorkDir.Length -gt 0 ) { $Shortcut.WorkingDirectory = $WorkDir }
				If($Icon.Length -gt 0){ $Shortcut.IconLocation = $Icon }
				$Shortcut.TargetPath = $TargetPath 
				$Shortcut.WindowStyle = $WindowStyle 
		 
		 
				Try 
					{ 
					    # Create Shortcut 
					    $Shortcut.Save() 
					    # Set Shortcut to Run Elevated 
					    If ($admin) 
							{      
								$TempFileName = [IO.Path]::GetRandomFileName() 
								$TempFile = [IO.FileInfo][IO.Path]::Combine($Path.Directory, $TempFileName) 
								$Writer = New-Object System.IO.FileStream $TempFile, ([System.IO.FileMode]::Create) 
								$Reader = $Path.OpenRead() 
								While ($Reader.Position -lt $Reader.Length) { 
									$Byte = $Reader.ReadByte() 
								    If ($Reader.Position -eq 22) {$Byte = 34} 
								    $Writer.WriteByte($Byte) 
								} 
								$Reader.Close() 
								$Writer.Close() 
								$Path.Delete() 
								Rename-Item -Path $TempFile -NewName $Path.Name | Out-Null 
					    	} 
					    Return $True 
				  	} 
				Catch 
					{ 
				    	Log-ScriptEvent -value "Unable to create $($Path.FullName)" -Component $CurrentScriptName -Severity 2 -OutToConsole $True
						Log-ScriptEvent -value $Error[0].Exception.Message -Component $CurrentScriptName -Severity 2 -OutToConsole $True  
				    	Return $False 
				  	} 
		 	} 
	}
	
	
	
Function Init-Function
	{

		param( 
		    	[Parameter(Mandatory=$false,Position=0)]  
		    	[string]$logPath
            )

        ##== Clean up $Error Cache
        $error.clear()

        ##== Gather OSD environement info into a global variable
		$Script:OSD_Env = Set-DeploymentEnv


        If ([string]::IsNullOrWhiteSpace($logPath)) 
            {
                If ($OSD_Env.BDDLog)
                    {
                        If (-not([string]::IsNullOrWhiteSpace($CurrentScriptPath)))
                            {$logPath = ((get-item $OSD_Env.BDDLog).Directoryname + "\" + $CurrentScriptName.replace("ps1","log"))}
                        Else
                            {$logPath = $($tsenv:Logpath +"\TestPSScript.log")}
                    }
                Else
                    {
                        If (-not([string]::IsNullOrWhiteSpace($CurrentScriptPath)))
                            {$logPath = ($CurrentScriptPath + "\" + $CurrentScriptName.replace("ps1","log"))}
                        Else
                            {$logPath = "C:\Temp\TestPSScript.log"}
                    }     
            }

        ##== Initialise Logging
		$Script:LogFile = Init-Logging -oPath $logPath

		If($OSD_Env.TSenv -eq $False) {Log-ScriptEvent "[WARNING] : unable to import MDT Powershell Module !!!" -Severity 2}

    	##== Gather Windows Info		
		$Script:OS_Env = Get-WindowsVersion
        #$Script:os_env.PSObject.Properties|foreach{Log-ScriptEvent $($_.Name + " : " + $_.Value)}
        #$Script:osd_env.PSObject.Properties|foreach{Log-ScriptEvent $($_.Name + " : " + $_.Value)}
		
		#Init Com object for aditional functions
		If ($OSD_Env.Tsenv -eq $true)
            {
                If (-not (Get-Process -name "TsProgressUI")){If($OSD_Env.IsSCCM -ne $true){Invoke-Executable -path "$($tsenv:DeployRoot)\Tools\x64\TsProgressUI.exe" -Arguments "/register"}}
                $Script:TaskSequenceProgressUi = New-Object -ComObject Microsoft.SMS.TSProgressUI
                    
            }
			
		$Script:SCCM_DevEnv = Set-SCCMDeveloppementEnv	
		
		If ($SCCM_DevEnv.SCCMDevModule -eq $true)
            {$Script:SCCM_Drive = Get-CMSite}
        Else 
		    {$Script:SCCM_Drive = $False}

		$Script:OSD_Env.psobject.properties|foreach {Log-ScriptEvent -value $('$Script:OSD_Env.' + $_.name + " = " +  $_.value)}
		$Script:OS_Env.psobject.properties|foreach {Log-ScriptEvent -value $('$Script:OS_Env.' + $_.name + " = " +  $_.value)}
		$Script:SCCM_DevEnv.psobject.properties|foreach {Log-ScriptEvent -value $('$Script:SCCM_DevEnv.' + $_.name + " = " +  $_.value)}
		$Script:SCCM_Drive.psobject.properties|foreach {Log-ScriptEvent -value $('$Script:SCCM_Drive.' + $_.name + " = " +  $_.value)}
		
		Log-ScriptEvent -value "Global Variables: LogFile, OSD_Env, OS_Env, SCCM_DevEnv and SCCM_Drive all set !! Function init finished"
		If (![String]::IsNullOrWhiteSpace($Script:TaskSequenceProgressUi))
			{Log-ScriptEvent "COM object Microsoft.SMS.TSProgressUI is register in variable `$TaskSequenceProgressUi"}
		Else
			{Log-ScriptEvent "[WARNING] Unable to register COM object Microsoft.SMS.TSProgressUI" -Severity 2}		
	}

	
function Show-TSActionProgress	
	{
	    <#
	    .SYNOPSIS
	    Shows task sequence secondary progress of a specific step
	    
	    .DESCRIPTION
	    Adds a second progress bar to the existing Task Sequence Progress UI.
	    This progress bar can be updated to allow for a real-time progress of
	    a specific task sequence sub-step.
	    The Step and Max Step parameters are calculated when passed. This allows
	    you to have a "max steps" of 400, and update the step parameter. 100%
	    would be achieved when step is 400 and max step is 400. The percentages
	    are calculated behind the scenes by the Com Object.
	    
	    .PARAMETER Message
	    The message to display the progress
	    .PARAMETER Step
	    Integer indicating current step
	    .PARAMETER MaxStep
	    Integer indicating 100%. A number other than 100 can be used.
	    .INPUTS
	     - Message: String
	     - Step: Long
	     - MaxStep: Long
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Set's "Custom Step 1" at 30 percent complete
	    Show-TSActionProgress -Message "Running Custom Step 1" -Step 100 -MaxStep 300
	    
	    .EXAMPLE
	    Set's "Custom Step 1" at 50 percent complete
	    Show-TSActionProgress -Message "Running Custom Step 1" -Step 150 -MaxStep 300
	    .EXAMPLE
	    Set's "Custom Step 1" at 100 percent complete
	    Show-TSActionProgress -Message "Running Custom Step 1" -Step 300 -MaxStep 300
	    #>
	    param(
	        [Parameter(Mandatory=$true)]
	        [string] $Message,
	        [Parameter(Mandatory=$true)]
	        [long] $Step,
	        [Parameter(Mandatory=$true)]
	        [long] $MaxStep
	    )

        If ($MaxStep -gt [uint32]::MaxValue)
            {
                [uint32]$MaxStep = ($MaxStep/10) 
                [uint32]$Step = ($Step/10)
            }

	    $TaskSequenceProgressUi.ShowActionProgress(`
	        $tsenv:_SMSTSOrgName,`
	        $tsenv:_SMSTSPackageName,`
	        $tsenv:_SMSTSCustomProgressDialogMessage,`
	        $tsenv:_SMSTSCurrentActionName,`
	        [Convert]::ToUInt32($tsenv:_SMSTSNextInstructionPointer),`
	        [Convert]::ToUInt32($tsenv:_SMSTSInstructionTableSize),`
	        $Message,`
	        $Step,`
	        $MaxStep)
	}


function Close-TSProgressDialog
	{
	    <#
	    .SYNOPSIS
	    Hides the Task Sequence Progress Dialog
	    
	    .DESCRIPTION
	    Hides the Task Sequence Progress Dialog
	    
	    .INPUTS
	    None
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Close-TSProgressDialog
	    #>

	    $TaskSequenceProgressUi.CloseProgressDialog()
	}

function Show-TSProgress
	{
	    <#
	    .SYNOPSIS
	    Shows task sequence progress of a specific step
	    
	    .DESCRIPTION
	    Manipulates the Task Sequence progress UI; top progress bar only.
	    This progress bar can be updated to allow for a real-time progress of
	    a specific task sequence step.
	    The Step and Max Step parameters are calculated when passed. This allows
	    you to have a "max steps" of 400, and update the step parameter. 100%
	    would be achieved when step is 400 and max step is 400. The percentages
	    are calculated behind the scenes by the Com Object.
	    
	    .PARAMETER CurrentAction
	    Step Title. Modifies the "Running action: " Message
	    .PARAMETER Step
	    Integer indicating current step
	    .PARAMETER MaxStep
	    Integer indicating 100%. A number other than 100 can be used.
	    .INPUTS
	     - CurrentAction: String
	     - Step: Long
	     - MaxStep: Long
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Set's "Custom Step 1" at 30 percent complete
	    Show-TSProgress -CurrentAction "Running Custom Step 1" -Step 100 -MaxStep 300
	    
	    .EXAMPLE
	    Set's "Custom Step 1" at 50 percent complete
	    Show-TSProgress -CurrentAction "Running Custom Step 1" -Step 150 -MaxStep 300
	    .EXAMPLE
	    Set's "Custom Step 1" at 100 percent complete
	    Show-TSProgress -CurrentAction "Running Custom Step 1" -Step 300 -MaxStep 300
	    #>
	    param(
	        [Parameter(Mandatory=$true)]
	        [string] $CurrentAction,
	        [Parameter(Mandatory=$true)]
	        [long] $Step,
	        [Parameter(Mandatory=$true)]
	        [long] $MaxStep
	    )

	    $TaskSequenceProgressUi.ShowTSProgress(`
	        $tsenv:_SMSTSOrgName, `
	        $tsenv:_SMSTSPackageName, `
	        $tsenv:_SMSTSCustomProgressDialogMessage, `
	        $CurrentAction, `
	        $Step, `
	        $MaxStep)

	}

function Show-TSErrorDialog
	{
		<#
	    .SYNOPSIS
	    Shows the Task Sequence Error Dialog
	    
	    .DESCRIPTION
	    Shows a task sequence error dialog allowing for custom failure pages.
	    
	    .PARAMETER OrganizationName
	    Name of your Organization
	    .PARAMETER CustomTitle
	    Custom Error Title
	    .PARAMETER ErrorMessage
	    Message details of the error
	    .PARAMETER ErrorCode
	    Error Code the Task sequence will exit with
	    .PARAMETER TimeoutInSeconds
	    Timout for the Reboot Prompt
	    .PARAMETER ForceReboot
	    Indicates whether a reboot will be forced or not
	    .INPUTS
	     - OrganizationName: String
	     - CustomTitle: String
	     - ErrorMessage: String
	     - ErrorCode: Long
	     - TimeoutInSeconds: Long
	     - ForceReboot: System.Boolean
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Sets an Error but does not force a reboot
	    Show-TSErrorDialog -OrganizationName "My Organization" -CustomTitle "An Error occured during the things" -ErrorMessage "That thing you tried...it didnt work" -ErrorCode 123456 -TimeoutInSeconds 90 -ForceReboot $false
	    
	    .EXAMPLE
	    Sets an Error and forces a reboot
	    Show-TSErrorDialog -OrganizationName "My Organization" -CustomTitle "An Error occured during the things" -ErrorMessage "He's dead Jim!" -ErrorCode 123456 -TimeoutInSeconds 90 -ForceReboot $true
	    #>
	    param(
	        [Parameter(Mandatory=$true)]
	        [string] $OrganizationName,
	        [Parameter(Mandatory=$true)]
	        [string] $CustomTitle,
	        [Parameter(Mandatory=$true)]
	        [string] $ErrorMessage,
	        [Parameter(Mandatory=$true)]
	        [long] $ErrorCode,
	        [Parameter(Mandatory=$true)]
	        [long] $TimeoutInSeconds,
	        [Parameter(Mandatory=$true)]
	        [bool] $ForceReboot
	    )

	    if ($ForceReboot)
		    {
		        $TaskSequenceProgressUi.ShowErrorDialog($OrganizationName, $Tsenv:_SMSTSPackageName, $CustomTitle, $ErrorMessage, $ErrorCode, $TimeoutInSeconds, 1)
		    }
	    else
		    {
		        $TaskSequenceProgressUi.ShowErrorDialog($OrganizationName, $Tsenv:_SMSTSPackageName, $CustomTitle, $ErrorMessage, $ErrorCode, $TimeoutInSeconds, 0)
		    }
	}

function Show-TSMessage
	{
	    <#
	    .SYNOPSIS
	    Shows a Windows Forms Message Box
	    
	    .DESCRIPTION
	    Shows a Windows Forms Message Box, but does not return the response.
	    This will halt any current operations while the prompt is shown.
	    
	    .PARAMETER Message
	    Message to be displayed
	    .PARAMETER Title
	    Title of the message box
	    .PARAMETER Type
	    Button Style for the MessageBox
	    0 = OK
	    1 = OK, Cancel
	    2 = Abort, Retry, Ignore
	    3 = Yes, No, Cancel
	    4 = Yes, No
	    5 = Retry, Cancel
	    6 = Cancel, Try Again, Continue
	    .INPUTS
	     - Message: String
	     - Title: String
	     - Type: Long
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Sets an Error but does not force a reboot
	    Show-TSErrorDialog -OrganizationName "My Organization" -CustomTitle "An Error occured during the things" -ErrorMessage "That thing you tried...it didnt work" -ErrorCode 123456 -TimeoutInSeconds 90 -ForceReboot $false
	    
	    .EXAMPLE
	    Sets an Error and forces a reboot
	    Show-TSErrorDialog -OrganizationName "My Organization" -CustomTitle "An Error occured during the things" -ErrorMessage "He's dead Jim!" -ErrorCode 123456 -TimeoutInSeconds 90 -ForceReboot $true
	    #>
	    param(
	        [Parameter(Mandatory=$true)]
	        [string] $Message,
	        [Parameter(Mandatory=$true)]
	        [string] $Title,
	        [Parameter(Mandatory=$true)]
	        [ValidateRange(0,6)]
	        [long] $Type
	    )

	    $TaskSequenceProgressUi.ShowMessage($Message, $Title, $Type)

	}

function Show-TSRebootDialog
	{
	    <#
	    .SYNOPSIS
	    Shows the Reboot Dialog
	    
	    .DESCRIPTION
	    Shows the Task Sequence "System Restart" Dialog. This allows you
	    to trigger custom Task Sequence Reboot Messages.
	    
	    .PARAMETER OrganizationName
	    Name of your Organization
	    .PARAMETER CustomTitle
	    Custom Title for the Reboot Dialog
	    .PARAMETER Message
	    Detailed Message regarding the reboot
	    .PARAMETER TimeoutInSeconds
	    Timout before the system reboots
	    .INPUTS
	     - OrganizationName: String
	     - CustomTitle: String
	     - Message: String
	     - TimeoutInSeconds: Long
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Show's a Reboot Dialog
	    Show-TSRebootDialog -OrganizationName "My Organization" -CustomTitle "I need a reboot!" -Message "I need to reboot to complete something..." -TimeoutInSeconds 90
	    #>
	    param(
	        [Parameter(Mandatory=$true)]
	        [string] $OrganizationName,
	        [Parameter(Mandatory=$true)]
	        [string] $CustomTitle,
	        [Parameter(Mandatory=$true)]
	        [string] $Message,
	        [Parameter(Mandatory=$true)]
	        [long] $TimeoutInSeconds
	    )

	    $TaskSequenceProgressUi.ShowRebootDialog($OrganizationName, $Tsenv:_SMSTSPackageName, $CustomTitle, $Message, $TimeoutInSeconds)
	}

function Show-TSSwapMediaDialog
	{
		<#
	    .SYNOPSIS
	    Shows Task Sequence Swap Media Dialog.
	    
	    .DESCRIPTION
	    Shows Task Sequence Swap Media Dialog.
	    
	    .PARAMETER TaskSequenceName
	    Name of the Task Sequence
	    .PARAMETER MediaNumber
	    Media Number to insert
	    .INPUTS
	     - TaskSequenceName: String
	     - CustomTitle: Long
	    .OUTPUTS
	    None
	    .EXAMPLE
	    Prompts to insert media #2 for the Task Sequence "My Task Sequence"
	    Show-TSSwapMediaDialog -TaskSequenceName "My Task Sequence" -MediaNumber 2
	    #>
	    param(
	        [Parameter(Mandatory=$true)]
	        [string] $TaskSequenceName,
	        [Parameter(Mandatory=$true)]
	        [long] $MediaNumber
	    )

	    $TaskSequenceProgressUi.ShowSwapMediaDialog($TaskSequenceName, $MediaNumber)

	}
	
	
function Invoke-Executable
	{
	   
		# usage:
		# $Iret = Invoke-Executable -Path "Setup.exe" -Arguments "/install /quiet /norestart"
		# The function return an array with the exit code in $Iret[0], the console output and the console ouput errors are starting at $Iret[1]
       		# Strings in returning valu can be found using -match :  $cmd = Invoke-Executable -path MyCommand.exe ; $cmd -match "*was successfull*"
        	# Added credencial support. Credencial must be provided with secure string (ex: using Get-credentical or )
		# Argument should be passed as a single string
		
		
		param(
	        [parameter(Mandatory=$true)]
	        [ValidateNotNullOrEmpty()]
	        [string]$Path,

	        [parameter(Mandatory=$false)]
	        [ValidateNotNull()]
                [string]$Arguments,
            
	        [parameter(Mandatory=$false)]
	        [ValidateNotNull()]
	        [SecureString]$Credential
	    )
		
		
		# Setup the Process startup info
		$pinfo = New-Object System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = $Path
		$pinfo.UseShellExecute = $false
		$pinfo.CreateNoWindow = $true
		$pinfo.RedirectStandardOutput = $true
		$pinfo.RedirectStandardError = $true
		
	        # Setup the Process argument
	        if (-not([String]::isnullorempty($Arguments))){$pinfo.Arguments = $Arguments}
        
	        # Setup the Process Credencial
	        if (-not([String]::isnullorempty($Credential)))
	            {
	                $pinfo.Username = $Credential.GetNetworkCredential().username
	                $pinfo.Domain = $Credential.GetNetworkCredential().Domain
	                $pinfo.Password = $Credential.Password
	            }    

		
		# Create a process object using the startup info
		$process = New-Object System.Diagnostics.Process
		Log-ScriptEvent -value "About to run command: $Path $Arguments"
		$process.StartInfo = $pinfo
		
				
	    # Invoke Start-Process cmdlet depending on if Arguments parameter input contains a object
        try 
			{$process.Start() | Out-Null}
        catch [System.Exception] 
			{$_ | Resolve-Error ; Break}
			
		while (!$process.HasExited){start-sleep -Seconds 1}
		Log-ScriptEvent -value "Process has existed with return code $($process.ExitCode)"
		
		# get output from stdout and stderr
		$stdout = $process.StandardOutput.ReadToEnd()
		$stderr = $process.StandardError.ReadToEnd()
        
        $Return = New-Object System.Collections.ArrayList
        $Null = $Return.add($process.ExitCode)
        
        Log-ScriptEvent -value "Returned console output:"
            
		if (-not([String]::IsNullOrEmpty($stdout))){foreach ($line in $stdout.split([Environment]::NewLine)){If(-not([String]::IsNullOrEmpty($line))){Log-ScriptEvent -value $line ; $Null = $Return.add($Line)}}}
		if (-not([String]::IsNullOrEmpty($stderr))){foreach ($line in $stderr.split([Environment]::NewLine)){If(-not([String]::IsNullOrEmpty($line))){Log-ScriptEvent -value $line ; $Null = $Return.add($Line)}}}
		
		Return $Return
	}	


Function Set-Registry
	{
	    <#
	    .SYNOPSIS
	    This function gives you the ability to create/change Windows registry keys and values. If you want to create a value but the key doesn't exist, it will create the key for you.
	    .PARAMETER RegKey
	    Path of the registry key to create/change
	    .PARAMETER RegValue
	    Name of the registry value to create/change
	    .PARAMETER RegData
	    The data of the registry value
	    .PARAMETER RegType
	    The type of the registry value. Allowed types: String,DWord,Binary,ExpandString,MultiString,None,QWord,Unknown. If no type is given, the function will use String as the type.
	    .EXAMPLE 
	    Set-Registry -RegKey HKLM:\SomeKey -RegValue SomeValue -RegData 1111 -RegType DWord
	    This will create the key SomeKey in HKLM:\. There it will create a value SomeValue of the type DWord with the data 1111.
	    .NOTES
	    Author: Dominik Britz
	    Source: https://github.com/DominikBritz
	    #>
	    [CmdletBinding()]
	    PARAM
	    (
	        $RegKey,
	        $RegValue,
	        $RegData,
	        [ValidateSet('String','DWord','Binary','ExpandString','MultiString','None','QWord','Unknown','REG_SZ','REG_MULTI_SZ','REG_EXPAND_SZ','REG_DWORD','REG_QWORD','REG_BINARY','REG_NONE')]
	        $RegType = 'String'    
	    )
		
		# Convert entries
		If ($RegType.toupper() -eq 'REG_SZ'){$RegType = 'String' }
		If ($RegType.toupper() -eq 'REG_MULTI_SZ'){$RegType = 'MultiString' }
		If ($RegType.toupper() -eq 'REG_EXPAND_SZ'){$RegType = 'ExpandString' }
		If ($RegType.toupper() -eq 'REG_DWORD'){$RegType = 'DWord' }
		If ($RegType.toupper() -eq 'REG_QWORD'){$RegType = 'QWord' }
		If ($RegType.toupper() -eq 'REG_BINARY'){$RegType = 'Binary' }
		If ($RegType.toupper() -eq 'REG_NONE'){$RegType = 'None' }
		
		If ($RegKey.Contains('HKEY_LOCAL_MACHINE')) {$RegKey = $RegKey.replace('HKEY_LOCAL_MACHINE',"HKLM:")}
		If ($RegKey.Contains('HKEY_USERS')) {$RegKey = $RegKey.replace('HKEY_USERS',"HKU:")}
		If ($RegKey.Contains('HKEY_CLASSES_ROOT')) {$RegKey = $RegKey.replace('HKEY_CLASSES_ROOT',"HKCR:")}
		If ($RegKey.Contains('HKEY_CURRENT_USER')) {$RegKey = $RegKey.replace('HKEY_CURRENT_USER',"HKCU:")}


	    If (-not $RegValue)
		    {
		        If (-not (Test-Path $RegKey))
			        {
			            Write-Verbose "The key $RegKey does not exist. Try to create it."
			            Try
				            {
				                New-Item -Path $RegKey -Force
				            }
			            Catch
				            {
				                Write-Error -Message $_
				            }
			            
						$Return = TestAndLog-RegistryValue -Path $RegKey -action "created"
			        }        
		    }

	    If ($RegValue)
		    {
		        If (-not (Test-Path $RegKey))
			        {
			            Write-Verbose "The key $RegKey does not exist. Try to create it."
			            Try
				            {
				                New-Item -Path $RegKey -Force
				                Set-ItemProperty -Path $RegKey -Name $RegValue -Value $RegData -Type $RegType -Force
				            }
			            Catch
				            {
				                Write-Error -Message $_
				            }
						$Return = TestAndLog-RegistryValue -Path $RegKey -Value $RegValue -action "created"
			        }
		        Else 
			        {
			            Write-Verbose "The key $RegKey already exists. Try to set value"
			            Try
				            {
				                Set-ItemProperty -Path $RegKey -Name $RegValue -Value $RegData -Type $RegType -Force
				            }
			            Catch
				            {
				                Write-Error -Message $_
				            }
						$Return = TestAndLog-RegistryValue -Path $RegKey -Value $RegValue -action "created"
			        }
		    }
		Return $Return	
	}


Function Import-MDTSnapin
	{
		If ((Get-Module |where name -eq "Microsoft.BDD.PSSnapIn") -or (Get-Module |where name -eq "MicrosoftDeploymentToolkit") )
			{
                		Log-ScriptEvent -value "MDT Powershell Module Detected"
                		Return $true
			}
		Else
			{
				IF (Test-Path 'C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1') 
					{	
						Log-ScriptEvent -value "Importing MDT Powershell SnapIn"
						Import-module 'C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
						Return $true
					}
				Else
					{
						Log-ScriptEvent -value "Unable to find file C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1, aborting !!!"
						Return $false				
					}
					
					
				If ((-not (Get-Module |where name -eq "Microsoft.BDD.PSSnapIn")) -or (-not(Get-Module |where name -eq "MicrosoftDeploymentToolkit")) ) 
					{
						Log-ScriptEvent -value "Unable to import MDT Module Microsoft.BDD.PSSnapIn, aborting !!!"
						Return $false			
					}
			}	
	}


function Dump-ObjectToXml
    {
        # This function convert an object to an xml file.
        # This function inspired by: https://msinnovations.wordpress.com/2011/07/29/export-powershell-object-to-xml-more-logical-than-explort-clixml/ 

        Param(
        	[parameter(Mandatory=$true)]
	        [ValidateNotNullOrEmpty()]
	        $obj,

	        [parameter(Mandatory=$false)]
	        [ValidateNotNull()]
	        [string]$xml,

	        [parameter(Mandatory=$false)]
	        [ValidateNotNull()]
	        [string]$tag
            )
        
        If (-not([String]::IsNullOrWhiteSpace($tag))) {$openingTag = "<" + $tag + ">"} else { $openingTag = "<" + $obj.GetType().Name + ">"}
        $ret = $ret + $openingTag
        Get-Member -InputObject $obj -MemberType Properties | ForEach-Object {
                $CurrentName = $_.Name
                $Val = $_.Definition.Substring($_.Definition.IndexOf("=")+1) # will catch everything after the first occurent of char '='
 
                $out = "<" + $_.Name + ">" + $val + "</" + $_.Name + ">"
                $ret = $ret + $out
             }
         
        If (-not([String]::IsNullOrWhiteSpace($tag))) {$closingTag = "</" + $tag + ">"} else {$closingTag = "</" + $a.GetType().Name + ">"}
        $ret = $ret + $closingTag
        If(-not([String]::IsNullOrWhiteSpace($xml))) 
			{
				$ret|Out-File $xml
				Log-ScriptEvent -value "Objects saved to file $xml"
			}
        Return $ret
    }


function Dump-XmlToObject
    {
       Param(
        	[parameter(Mandatory=$true)]
	        [ValidateNotNull()]
	        [string]$xml 
            )

        If (test-path $xml) 
            {[xml]$xmlFile = Get-Content $xml} 
        Else 
            {Log-ScriptEvent -value "[ERROR] Unable to upload MDT xml file, aborting !!!" ; Return $false}

        $obj = New-Object System.Object
        $xmlSelect = $xmlfile.SelectNodes("//*")

        foreach ($item in $xmlSelect){If (-not([String]::IsNullOrWhiteSpace(($item.'#text')))){$obj|Add-Member -MemberType NoteProperty -Name $item.Name -Value $item.'#text'}}
        Return $obj
    }
	
	
Function Add-IniComment {
    <#
    .Synopsis
        Comments out specified content of an INI file
    .Description
        Comments out specified keys in all sections or certain sections.
        The ini source can be specified by a file or piped in by the result of Get-IniContent.
        The modified content is returned as a ordered dictionary hashtable and can be piped to a file with Out-IniFile.
    .Notes
        Author		: Sean Seymour <seanjseymour@gmail.com> based on work by Oliver Lipkau <oliver@lipkau.net>
		Source		: https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version		: 1.0.0 - 2016/08/18 - SS - Initial release
                    : 1.0.1 - 2016/12/29 - SS - Removed need for delimiters by making Sections and Keys string arrays.
        #Requires -Version 2.0
    .Inputs
        System.String
        System.Collections.IDictionary
    .Outputs
        System.Collections.Specialized.OrderedDictionary
    .Example
        $ini = Add-IniComment -FilePath "C:\myinifile.ini" -Sections 'Printers' -Keys 'Headers','Footers'
        -----------
        Description
        Reads in the INI File c:\myinifile.ini, comments out any keys named 'Headers' or 'Footers' in the [Printers] section, and saves the modified ini to $ini.
    .Example
        Add-IniComment -FilePath "C:\myinifile.ini" -Sections 'Terminals','Monitors' -Keys 'Updated' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini and comments out any keys named 'Updated' in the [Terminals] and [Monitors] sections.
        The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Add-IniComment -Keys 'Headers' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Add-IniComment to comment out any 'Headers' keys in any
        section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Add-IniComment -Keys 'Updated' -Sections '_' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Add-IniComment to comment out any 'Updated' keys that
        are orphaned, i.e. not specifically in a section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini.
    .Link
        Get-IniContent
        Out-IniFile
    #>

    [CmdletBinding(DefaultParameterSetName = "File")]
    [OutputType(
        [System.Collections.Specialized.OrderedDictionary]
    )]
    Param
    (
        # Specifies the path to the input file.
        [Parameter( Position = 0, Mandatory = $true, ParameterSetName = "File" )]
        [ValidateNotNullOrEmpty()]
        [String]
        $FilePath,

        # Specifies the Hashtable to be modified. Enter a variable that contains the objects or type a command or expression that gets the objects.
        [Parameter( Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Object" )]
        [ValidateNotNullOrEmpty()]
        [System.Collections.IDictionary]
        $InputObject,

        # String array of one or more keys to limit the changes to, separated by a comma. Optional.
        [Parameter( Mandatory = $true )]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Keys,

        # Specify what character should be used to comment out entries.
        # Note: This parameter is a char array to maintain compatibility with the other functions.
        # However, only the first character is used to comment out entries.
        # Default: ";"
        [Char[]]
        $CommentChar = @(";"),

        # String array of one or more sections to limit the changes to, separated by a comma.
        # Surrounding section names with square brackets is not necessary but is supported.
        # Ini keys that do not have a defined section can be modified by specifying '_' (underscore) for the section.
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Sections
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"

        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"
    }

    Process {
        # Get the ini from either a file or object passed in.
        if ($PSCmdlet.ParameterSetName -eq 'File') { $content = Get-IniContent $FilePath }
        if ($PSCmdlet.ParameterSetName -eq 'Object') { $content = $InputObject }

        # Specific section(s) were requested.
        if ($Sections) {
            foreach ($section in $Sections) {
                # Get rid of whitespace and section brackets.
                $section = $section.Trim() -replace '[][]', ''

                Write-Debug ("Processing '{0}' section." -f $section)

                foreach ($key in $Keys) {
                    Write-Debug ("Processing '{0}' key." -f $key)

                    $key = $key.Trim()

                    if ($content[$section]) {
                        $currentValue = $content[$section][$key]
                    }
                    else {
                        Write-Verbose ("$($MyInvocation.MyCommand.Name):: '{0}' section does not exist." -f $section)
                        # Break out of the loop after this, because we don't want to check further keys for this non-existent section.
                        break
                    }

                    if ($currentValue) {
                        Convert-IniEntryToComment $content $key $section $CommentChar
                    }
                    else {
                        Write-Verbose ("$($MyInvocation.MyCommand.Name):: '[{0}][{1}]' does not exist." -f $section, $key)
                    }
                }
            }
        }
        else {
            # No section supplied, go through the entire ini since changes apply to all sections.
            foreach ($item in $content.GetEnumerator()) {
                $section = $item.key
                Write-Debug ("Processing '{0}' section." -f $section)

                foreach ($key in $Keys) {
                    $key = $key.Trim()
                    Write-Debug ("Processing '{0}' key." -f $key)

                    if ($content[$section][$key]) {
                        Convert-IniEntryToComment $content $key $section $CommentChar
                    }
                    else {
                        Write-Verbose ("$($MyInvocation.MyCommand.Name):: '[{0}][{1}]' does not exist." -f $section, $key)
                    }
                }
            }
        }

        Write-Output $content
    }
    End {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}

Function Convert-IniCommentToEntry {
    <#
    .SYNOPSIS
        Internal module function to remove the old comment then insert a new key/value pair at the old location with the previous comment's value.
    #>
    param ($content, $key, $section, $commentChar)

    $index = 0
    $commentFound = $false

    $commentRegex = "^([$($commentChar -join '')]$key.*)$"
    Write-Debug ("commentRegex is {0}." -f $commentRegex)

    foreach ($entry in $content[$section].GetEnumerator()) {
        Write-Debug ("Uncomment looking at key '{0}' with value '{1}'." -f $entry.key, $entry.value)

        if ($entry.key.StartsWith('Comment') -and $entry.value -match $commentRegex) {
            Write-Verbose ("$($MyInvocation.MyCommand.Name):: Uncommenting '{0}' in {1} section." -f $entry.value, $section)
            $oldKey = $entry.key
            $split = $entry.value.Split("=")

            if ($split.Length -ge 2) {
                $newValue = $split[1].Trim()
            }
            else {
                # If the split did not result in 2+ items, it was not in the key=value form.
                # So just uncomment the key, as there is no value. It will result in a "key=" formatted output.
                $newValue = ''
            }

            # Break out once a match is found. If there are multiple commented out keys
            # with the same name, we can't add them anyway since it's a hash.
            $commentFound = $true
            break
        }
        $index++
    }

    if ($commentFound) {
        if ($content[$section][$key]) {
            Write-Verbose ("$($MyInvocation.MyCommand.Name):: Unable to uncomment '{0}' key in {1} section as there is already a key with that name." -f $key, $section)
        }
        else {
            Write-Debug ("Removing '{0}'." -f $oldKey)
            $content[$section].Remove($oldKey)
            Write-Debug ("Inserting [{0}][{1}] = {2} at index {3}." -f $section, $key, $newValue, $index)
            $content[$section].Insert($index, $key, $newValue)
        }
    }
    else {
        Write-Verbose ("$($MyInvocation.MyCommand.Name):: Did not find '{0}' key in {1} section to uncomment." -f $key, $section)
    }
}

Function Convert-IniEntryToComment {
    <#
    .SYNOPSIS
        Internal module function to remove the old key then insert a new one at the old location in the comment style used by Get-IniContent.
    #>
    param ($content, $key, $section, $commentChar)

    # Comments in Get-IniContent start with 1, not zero.
    $commentCount = 1

    foreach ($entry in $content[$section].GetEnumerator()) {
        if ($entry.key.StartsWith('Comment')) {
            $commentCount++
        }
    }

    Write-Debug ("commentCount is {0}." -f $commentCount)

    $desiredValue = $content[$section][$key]

    # Don't attempt to comment out non-existent keys.
    if ($desiredValue) {
        Write-Debug ("desiredValue is {0}." -f $desiredValue)

        $commentKey = 'Comment' + $commentCount
        Write-Debug ("commentKey is {0}." -f $commentKey)

        $commentValue = $commentChar[0] + $key + '=' + $desiredValue
        Write-Debug ("commentValue is {0}." -f $commentValue)

        # Thanks to http://stackoverflow.com/a/35731603/844937. However, that solution is case sensitive.
        # Tried $index = $($content[$section].keys).IndexOf($key, [StringComparison]"CurrentCultureIgnoreCase")
        # but it said there were no IndexOf overloads with two arguments. So if we get a -1 (not found),
        # use a variation on http://stackoverflow.com/a/34930231/844937 to search for a case-insensitive match.
        $sectionKeys = $($content[$section].keys)
        $index = $sectionKeys.IndexOf($key)
        Write-Debug ("Index of {0} is {1}." -f $key, $index)

        if ($index -eq -1) {
            $i = 0
            foreach ($sectionKey in $sectionKeys) {
                if ($sectionKey -match $key) {
                    $index = $i
                    Write-Debug ("Index updated to {0}." -f $index)
                    break
                }
                else {
                    $i++
                }
            }
        }

        if ($index -ge 0) {
            Write-Verbose ("$($MyInvocation.MyCommand.Name):: Commenting out {0} key in {1} section." -f $key, $section)
            $content[$section].Remove($key)
            $content[$section].Insert($index, $commentKey, $commentValue)
        }
        else {
            Write-Verbose ("$($MyInvocation.MyCommand.Name):: Could not find '{0}' key in {1} section to comment out." -f $key, $section)
        }
    }
}


Function Get-IniContent {
    <#
    .Synopsis
        Gets the content of an INI file
    .Description
        Gets the content of an INI file and returns it as a hashtable
    .Notes
        Author		: Oliver Lipkau <oliver@lipkau.net>
		Source		: https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version		: 1.0.0 - 2010/03/12 - OL - Initial release
                      1.0.1 - 2014/12/11 - OL - Typo (Thx SLDR)
                                              Typo (Thx Dave Stiff)
                      1.0.2 - 2015/06/06 - OL - Improvment to switch (Thx Tallandtree)
                      1.0.3 - 2015/06/18 - OL - Migrate to semantic versioning (GitHub issue#4)
                      1.0.4 - 2015/06/18 - OL - Remove check for .ini extension (GitHub Issue#6)
                      1.1.0 - 2015/07/14 - CB - Improve round-tripping and be a bit more liberal (GitHub Pull #7)
                                           OL - Small Improvments and cleanup
                      1.1.1 - 2015/07/14 - CB - changed .outputs section to be OrderedDictionary
                      1.1.2 - 2016/08/18 - SS - Add some more verbose outputs as the ini is parsed,
                      				            allow non-existent paths for new ini handling,
                      				            test for variable existence using local scope,
                      				            added additional debug output.
        #Requires -Version 2.0
    .Inputs
        System.String
    .Outputs
        System.Collections.Specialized.OrderedDictionary
    .Example
        $FileContent = Get-IniContent "C:\myinifile.ini"
        -----------
        Description
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent
    .Example
        $inifilepath | $FileContent = Get-IniContent
        -----------
        Description
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent
    .Example
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini"
        C:\PS>$FileContent["Section"]["Key"]
        -----------
        Description
        Returns the key "Key" of the section "Section" from the C:\settings.ini file
    .Link
        Out-IniFile
    #>

    [CmdletBinding()]
    [OutputType(
        [System.Collections.Specialized.OrderedDictionary]
    )]
    Param(
        # Specifies the path to the input file.
        [ValidateNotNullOrEmpty()]
        [Parameter( Mandatory = $true, ValueFromPipeline = $true )]
        [String]
        $FilePath,

        # Specify what characters should be describe a comment.
        # Lines starting with the characters provided will be rendered as comments.
        # Default: ";"
        [Char[]]
        $CommentChar = @(";"),

        # Remove lines determined to be comments from the resulting dictionary.
        [Switch]
        $IgnoreComments
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"

        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"

        $commentRegex = "^([$($CommentChar -join '')].*)$"

        Write-Debug ("commentRegex is {0}." -f $commentRegex)
    }

    Process {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath"

        $ini = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)

        if (!(Test-Path $Filepath)) {
            Write-Verbose ("Warning: `"{0}`" was not found." -f $Filepath)
            Write-Output $ini
        }

        $commentCount = 0
        switch -regex -file $FilePath {
            "^\s*\[(.+)\]\s*$" {
                # Section
                $section = $matches[1]
                Write-Verbose "$($MyInvocation.MyCommand.Name):: Adding section : $section"
                $ini[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                $CommentCount = 0
                continue
            }
            $commentRegex {
                # Comment
                if (!$IgnoreComments) {
                    if (!(test-path "variable:local:section")) {
                        $section = $script:NoSection
                        $ini[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                    }
                    $value = $matches[1]
                    $CommentCount++
                    Write-Debug ("Incremented CommentCount is now {0}." -f $CommentCount)
                    $name = "Comment" + $CommentCount
                    Write-Verbose "$($MyInvocation.MyCommand.Name):: Adding $name with value: $value"
                    $ini[$section][$name] = $value
                }
                else {
                    Write-Debug ("Ignoring comment {0}." -f $matches[1])
                }

                continue
            }
            "(.+?)\s*=\s*(.*)" {
                # Key
                if (!(test-path "variable:local:section")) {
                    $section = $script:NoSection
                    $ini[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                }
                $name, $value = $matches[1..2]
                Write-Verbose "$($MyInvocation.MyCommand.Name):: Adding key $name with value: $value"
                $ini[$section][$name] = $value
                continue
            }
        }
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"
        Write-Output $ini
    }

    End {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}


Function Out-IniFile {
    <#
    .Synopsis
        Write hash content to INI file
    .Description
        Write hash content to INI file
    .Notes
        Author      : Oliver Lipkau <oliver@lipkau.net>
        Blog        : http://oliver.lipkau.net/blog/
        Source      : https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        #Requires -Version 2.0
    .Inputs
        System.String
        System.Collections.IDictionary
    .Outputs
        System.IO.FileSystemInfo
    .Example
        Out-IniFile $IniVar "C:\myinifile.ini"
        -----------
        Description
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini
    .Example
        $IniVar | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini and overwrites the file if it is already present
    .Example
        $file = Out-IniFile $IniVar "C:\myinifile.ini" -PassThru
        -----------
        Description
        Saves the content of the $IniVar Hashtable to the INI File c:\myinifile.ini and saves the file into $file
    .Example
        $Category1 = @{“Key1”=”Value1”;”Key2”=”Value2”}
        $Category2 = @{“Key1”=”Value1”;”Key2”=”Value2”}
        $NewINIContent = @{“Category1”=$Category1;”Category2”=$Category2}
        Out-IniFile -InputObject $NewINIContent -FilePath "C:\MyNewFile.ini"
        -----------
        Description
        Creating a custom Hashtable and saving it to C:\MyNewFile.ini
    .Link
        Get-IniContent
    #>

    [CmdletBinding()]
    [OutputType(
        [System.IO.FileSystemInfo]
    )]
    Param(
        # Adds the output to the end of an existing file, instead of replacing the file contents.
        [switch]
        $Append,

        # Specifies the file encoding. The default is UTF8.
        #
        # Valid values are:
        # -- ASCII:  Uses the encoding for the ASCII (7-bit) character set.
        # -- BigEndianUnicode:  Encodes in UTF-16 format using the big-endian byte order.
        # -- Byte:   Encodes a set of characters into a sequence of bytes.
        # -- String:  Uses the encoding type for a string.
        # -- Unicode:  Encodes in UTF-16 format using the little-endian byte order.
        # -- UTF7:   Encodes in UTF-7 format.
        # -- UTF8:  Encodes in UTF-8 format.
        [ValidateSet("Unicode", "UTF7", "UTF8", "ASCII", "BigEndianUnicode", "Byte", "String")]
        [Parameter()]
        [String]
        $Encoding = "UTF8",

        # Specifies the path to the output file.
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {Test-Path $_ -IsValid} )]
        [Parameter( Position = 0, Mandatory = $true )]
        [String]
        $FilePath,

        # Allows the cmdlet to overwrite an existing read-only file. Even using the Force parameter, the cmdlet cannot override security restrictions.
        [Switch]
        $Force,

        # Specifies the Hashtable to be written to the file. Enter a variable that contains the objects or type a command or expression that gets the objects.
        [Parameter( Mandatory = $true, ValueFromPipeline = $true )]
        [System.Collections.IDictionary]
        $InputObject,

        # Passes an object representing the location to the pipeline. By default, this cmdlet does not generate any output.
        [Switch]
        $Passthru,

        # Adds spaces around the equal sign when writing the key = value
        [Switch]
        $Loose,

        # Writes the file as "pretty" as possible
        #
        # Adds an extra linebreak between Sections
        [Switch]
        $Pretty
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"

        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"

        function Out-Keys {
            param(
                [ValidateNotNullOrEmpty()]
                [Parameter( Mandatory, ValueFromPipeline )]
                [System.Collections.IDictionary]
                $InputObject,

                [ValidateSet("Unicode", "UTF7", "UTF8", "ASCII", "BigEndianUnicode", "Byte", "String")]
                [Parameter( Mandatory )]
                [string]
                $Encoding = "UTF8",

                [ValidateNotNullOrEmpty()]
                [ValidateScript( {Test-Path $_ -IsValid})]
                [Parameter( Mandatory, ValueFromPipelineByPropertyName )]
                [string]
                $Path,

                [Parameter( Mandatory )]
                $Delimiter,

                [Parameter( Mandatory )]
                $MyInvocation
            )

            Process {
                if (!($InputObject.get_keys())) {
                    Write-Warning ("No data found in '{0}'." -f $FilePath)
                }
                Foreach ($key in $InputObject.get_keys()) {
                    if ($key -match "^Comment\d+") {
                        Write-Verbose "$($MyInvocation.MyCommand.Name):: Writing comment: $key"
                        Add-Content -Value "$($InputObject[$key])" -Encoding $Encoding -Path $Path
                    }
                    else {
                        Write-Verbose "$($MyInvocation.MyCommand.Name):: Writing key: $key"
                        Add-Content -Value "$key$delimiter$($InputObject[$key])" -Encoding $Encoding -Path $Path
                    }
                }
            }
        }

        $delimiter = '='
        if ($Loose) {
            $delimiter = ' = '
        }

        # Splatting Parameters
        $parameters = @{
            Encoding = $Encoding;
            Path     = $FilePath
        }

    }

    Process {
        $extraLF = ""

        if ($Append) {
            Write-Debug ("Appending to '{0}'." -f $FilePath)
            $outfile = Get-Item $FilePath
        }
        else {
            Write-Debug ("Creating new file '{0}'." -f $FilePath)
            $outFile = New-Item -ItemType file -Path $Filepath -Force:$Force
        }

        if (!(Test-Path $outFile.FullName)) {Throw "Could not create File"}

        Write-Verbose "$($MyInvocation.MyCommand.Name):: Writing to file: $Filepath"
        foreach ($i in $InputObject.get_keys()) {
            if (!($InputObject[$i].GetType().GetInterface('IDictionary'))) {
                #Key value pair
                Write-Verbose "$($MyInvocation.MyCommand.Name):: Writing key: $i"
                Add-Content -Value "$i$delimiter$($InputObject[$i])" @parameters

            }
            elseif ($i -eq $script:NoSection) {
                #Key value pair of NoSection
                Out-Keys $InputObject[$i] `
                    @parameters `
                    -Delimiter $delimiter `
                    -MyInvocation $MyInvocation
            }
            else {
                #Sections
                Write-Verbose "$($MyInvocation.MyCommand.Name):: Writing Section: [$i]"

                # Only write section, if it is not a dummy ($script:NoSection)
                if ($i -ne $script:NoSection) { Add-Content -Value "$extraLF[$i]" @parameters }
                if ($Pretty) {
                    $extraLF = "`r`n"
                }

                if ( $InputObject[$i].Count) {
                    Out-Keys $InputObject[$i] `
                        @parameters `
                        -Delimiter $delimiter `
                        -MyInvocation $MyInvocation
                }

            }
        }
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Writing to file: $FilePath"
    }

    End {
        if ($PassThru) {
            Write-Debug ("Returning file due to PassThru argument.")
            Write-Output (Get-Item $outFile)
        }
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}



Function Remove-IniComment {
    <#
    .Synopsis
        Uncomments out specified content of an INI file
    .Description
        Uncomments out specified keys in all sections or certain sections.
        The ini source can be specified by a file or piped in by the result of Get-IniContent.
        The modified content is returned as a ordered dictionary hashtable and can be piped to a file with Out-IniFile.
    .Notes
        Author		: Sean Seymour <seanjseymour@gmail.com> based on work by Oliver Lipkau <oliver@lipkau.net>
		Source		: https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version		: 1.0.0 - 2016/08/18 - SS - Initial release
                    : 1.0.1 - 2016/12/29 - SS - Removed need for delimiters by making Sections and Keys string arrays.
        #Requires -Version 2.0
    .Inputs
        System.String
        System.Collections.IDictionary
    .Outputs
        System.Collections.Specialized.OrderedDictionary
    .Example
        $ini = Remove-IniComment -FilePath "C:\myinifile.ini" -Sections 'Printers' -Keys 'Headers'
        -----------
        Description
        Reads in the INI File c:\myinifile.ini, uncomments out any keys named 'Headers' in the [Printers] section, and saves the modified ini to $ini.
    .Example
        Remove-IniComment -FilePath "C:\myinifile.ini" -Sections 'Terminals','Monitors' -Keys 'Updated' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini and uncomments out any keys named 'Updated' in the [Terminals] and [Monitors] sections.
        The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Remove-IniComment -Keys 'Headers' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Remove-IniComment to uncomment any 'Headers' keys in any
        section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Remove-IniComment -Keys 'Updated' -Sections '_' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Remove-IniComment to uncomment any 'Updated' keys that
        are orphaned, i.e. not specifically in a section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini.
    .Link
        Get-IniContent
        Out-IniFile
    #>

    [CmdletBinding( DefaultParameterSetName = "File" )]
    [OutputType(
        [System.Collections.IDictionary]
    )]
    Param
    (
        # Specifies the path to the input file.
        [Parameter( Position = 0,  Mandatory = $true, ParameterSetName = "File" )]
        [ValidateNotNullOrEmpty()]
        [String]
        $FilePath,

        # Specifies the Hashtable to be modified. Enter a variable that contains the objects or type a command or expression that gets the objects.
        [Parameter( Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Object" )]
        [ValidateNotNullOrEmpty()]
        [System.Collections.IDictionary]
        $InputObject,

        # String array of one or more keys to limit the changes to, separated by a comma. Optional.
        [Parameter( Mandatory = $true )]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Keys,

        # Specify what characters should be describe a comment.
        # Lines starting with the characters provided will be rendered as comments.
        # Default: ";"
        [Char[]]
        $CommentChar = @(";"),

        # String array of one or more sections to limit the changes to, separated by a comma.
        # Surrounding section names with square brackets is not necessary but is supported.
        # Ini keys that do not have a defined section can be modified by specifying '_' (underscore) for the section.
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Sections
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"

        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"
    }
    # Uncomment out the specified keys in the list, either in the specified section or in all sections.
    Process {
        # Get the ini from either a file or object passed in.
        if ($PSCmdlet.ParameterSetName -eq 'File') { $content = Get-IniContent $FilePath }
        if ($PSCmdlet.ParameterSetName -eq 'Object') { $content = $InputObject }

        # Specific section(s) were requested.
        if ($Sections) {
            foreach ($section in $Sections) {
                # Get rid of whitespace and section brackets.
                $section = $section.Trim() -replace '[][]', ''

                Write-Debug ("Processing '{0}' section." -f $section)

                foreach ($key in $Keys) {
                    Write-Debug ("Processing '{0}' key." -f $key)

                    $key = $key.Trim()

                    if (!($content[$section])) {
                        Write-Verbose ("$($MyInvocation.MyCommand.Name):: '{0}' section does not exist." -f $section)
                        # Break out of the loop after this, because we don't want to check further keys for this non-existent section.
                        break
                    }
                    # Since this is a comment, we need to search through all the CommentX keys in this section.
                    # That's handled in the Convert-IniCommentToEntry function, so don't bother checking key existence here.
                    Convert-IniCommentToEntry $content $key $section $CommentChar
                }
            }
        }
        else {
            # No section supplied, go through the entire ini since changes apply to all sections.
            foreach ($item in $content.GetEnumerator()) {
                $section = $item.key
                Write-Debug ("Processing '{0}' section." -f $section)

                foreach ($key in $Keys) {
                    $key = $key.Trim()
                    Write-Debug ("Processing '{0}' key." -f $key)
                    Convert-IniCommentToEntry $content $key $section $CommentChar
                }
            }
        }

        Write-Output $content
    }
    End {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}


Function Remove-IniEntry {
    <#
    .Synopsis
        Removes specified content from an INI file
    .Description
        Removes specified keys in all sections or certain sections.
        The ini source can be specified by a file or piped in by the result of Get-IniContent.
        The modified content is returned as a ordered dictionary hashtable and can be piped to a file with Out-IniFile.
    .Notes
        Author		: Sean Seymour <seanjseymour@gmail.com> based on work by Oliver Lipkau <oliver@lipkau.net>
		Source		: https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version		: 1.0.0 - 2016/08/18 - SS - Initial release
                    : 1.0.1 - 2016/12/29 - SS - Removed need for delimiters by making Sections and Keys string arrays.
        #Requires -Version 2.0
    .Inputs
        System.String
        System.Collections.IDictionary
    .Outputs
        System.Collections.Specialized.OrderedDictionary
    .Example
        $ini = Remove-IniEntry -FilePath "C:\myinifile.ini" -Sections 'Printers' -Keys 'Headers','Version'
        -----------
        Description
        Reads in the INI File c:\myinifile.ini, removes any keys named 'Headers' or 'Version' in the [Printers] section, and saves the modified ini to $ini.
    .Example
        Remove-IniEntry -FilePath "C:\myinifile.ini" -Sections 'Terminals','Monitors' -Keys 'Updated' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini and removes any keys named 'Updated' in the [Terminals] and [Monitors] sections.
        The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Remove-IniEntry -Keys 'Headers' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Remove-IniEntry to remove any 'Headers' keys in any
        section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Remove-IniEntry -Sections 'Terminals' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Remove-IniEntry to remove the 'Terminals' section.
        The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Remove-IniEntry -Keys 'Updated' -Sections '_' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Remove-IniEntry to remove any 'Updated' keys that
        are orphaned, i.e. not specifically in a section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini.
    .Link
        Get-IniContent
        Out-IniFile
    #>

    [CmdletBinding(DefaultParameterSetName = "File")]
    [OutputType(
        [System.Collections.IDictionary]
    )]
    Param
    (
        # Specifies the path to the input file.
        [Parameter( Position = 0, Mandatory = $true, ParameterSetName = "File")]
        [ValidateNotNullOrEmpty()]
        [String]
        $FilePath,

        # Specifies the Hashtable to be modified.
        # Enter a variable that contains the objects or type a command or expression that gets the objects.
        [Parameter( Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Object" )]
        [System.Collections.IDictionary]
        $InputObject,

        # String array of one or more keys to limit the changes to, separated by a comma. Optional.
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Keys,

        [ValidateNotNullOrEmpty()]
        [String[]]
        $Sections
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"
    }
    # Remove the specified keys in the list, either in the specified section or in all sections.
    Process {
        # Get the ini from either a file or object passed in.
        if ($PSCmdlet.ParameterSetName -eq 'File') { $content = Get-IniContent $FilePath }
        if ($PSCmdlet.ParameterSetName -eq 'Object') { $content = $InputObject }

        if (!$Keys -and !$Sections) {
            Write-Verbose ("No sections or keys provided, exiting.")
            Write-Output $content
        }

        # Specific section(s) were requested.
        if ($Sections) {
            foreach ($section in $Sections) {
                # Get rid of whitespace and section brackets.
                $section = $section.Trim() -replace '[][]', ''

                Write-Debug ("Processing '{0}' section." -f $section)

                # If the user wants to remove an entire section, there will be a section specified but no keys.
                if (!$Keys) {
                    Write-Verbose ("Deleting entire section '{0}'." -f $section)
                    $content.Remove($section)
                }
                else {
                    foreach ($key in $Keys) {
                        Write-Debug ("Processing '{0}' key." -f $key)

                        $key = $key.Trim()

                        if ($content[$section]) {
                            $currentValue = $content[$section][$key]
                        }
                        else {
                            Write-Verbose ("$($MyInvocation.MyCommand.Name):: '{0}' section does not exist." -f $section)
                            # Break out of the loop after this, because we don't want to check further keys for this non-existent section.
                            break
                        }

                        if ($currentValue) {
                            Write-Verbose ("Removing {0} key from {1} section." -f $key, $section)
                            $content[$section].Remove($key)
                        }
                        else {
                            Write-Verbose ("$($MyInvocation.MyCommand.Name):: '{0}' key does not exist." -f $key)
                        }
                    }
                }
            }
        }
        else {
            # No section supplied, go through the entire ini since changes apply to all sections.
            foreach ($item in $content.GetEnumerator()) {
                $section = $item.key
                Write-Debug ("Processing '{0}' section." -f $section)

                foreach ($key in $Keys) {
                    $key = $key.Trim()
                    Write-Debug ("Processing '{0}' key." -f $key)

                    if ($content[$section][$key]) {
                        Write-Verbose ("Removing {0} key from {1} section." -f $key, $section)
                        $content[$section].Remove($key)
                    }
                    else {
                        Write-Verbose ("$($MyInvocation.MyCommand.Name):: '{0}' key does not exist in {1} section." -f $key, $section)
                    }
                }
            }
        }

        Write-Output $content
    }
    End {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}

Function Set-IniContent {
    <#
    .Synopsis
        Updates existing values or adds new key-value pairs to an INI file
    .Description
        Updates specified keys to new values in all sections or certain sections.
        Used to add new or change existing values. To comment, uncomment or remove keys use the related functions instead.
        The ini source can be specified by a file or piped in by the result of Get-IniContent.
        The modified content is returned as a ordered dictionary hashtable and can be piped to a file with Out-IniFile.
    .Notes
        Author		: Sean Seymour <seanjseymour@gmail.com> based on work by Oliver Lipkau <oliver@lipkau.net>
		Source		: https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version		: 1.0.0 - 2016/08/18 - SS - Initial release
                    : 1.0.1 - 2016/12/29 - SS - Removed need for delimiters by making Sections a string array
                                                and NameValuePairs a hashtable. Thanks Oliver!
        #Requires -Version 2.0
    .Inputs
        System.String
        System.Collections.IDictionary
    .Outputs
        System.Collections.Specialized.OrderedDictionary
    .Example
        $ini = Set-IniContent -FilePath "C:\myinifile.ini" -Sections 'Printers' -NameValuePairs @{'Name With Space' = 'Value1' ; 'AnotherName' = 'Value2'}
        -----------
        Description
        Reads in the INI File c:\myinifile.ini, adds or updates the 'Name With Space' and 'AnotherName' keys in the [Printers] section to the values specified,
        and saves the modified ini to $ini.
    .Example
        Set-IniContent -FilePath "C:\myinifile.ini" -Sections 'Terminals','Monitors' -NameValuePairs @{'Updated=FY17Q2'} | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini and adds or updates the 'Updated' key in the [Terminals] and [Monitors] sections to the value specified.
        The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Set-IniContent -NameValuePairs @{'Headers' = 'True' ; 'Update' = 'False'} | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Set-IniContent to add or update the 'Headers'  and 'Update' keys in all sections
        to the specified values. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini. If the file is already present it will be overwritten.
    .Example
        Get-IniContent "C:\myinifile.ini" | Set-IniContent -NameValuePairs @{'Updated'='FY17Q2'} -Sections '_' | Out-IniFile "C:\myinifile.ini" -Force
        -----------
        Description
        Reads in the INI File c:\myinifile.ini using Get-IniContent, which is then piped to Set-IniContent to add or update the 'Updated' key that
        is orphaned, i.e. not specifically in a section. The ini is then piped to Out-IniFile to write the INI File to c:\myinifile.ini.
    .Link
        Get-IniContent
        Out-IniFile
    #>

    [CmdletBinding(DefaultParameterSetName = "File")]
    [OutputType(
        [System.Collections.IDictionary]
    )]
    Param
    (
        # Specifies the path to the input file.
        [Parameter( Position = 0, Mandatory = $true, ParameterSetName = "File" )]
        [ValidateNotNullOrEmpty()]
        [String]
        $FilePath,

        # Specifies the Hashtable to be modified.
        # Enter a variable that contains the objects or type a command or expression that gets the objects.
        [Parameter( Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Object")]
        [ValidateNotNullOrEmpty()]
        [System.Collections.IDictionary]
        $InputObject,

        # Hashtable of one or more key names and values to modify. Required.
        [Parameter( Mandatory = $true, ParameterSetName = "File")]
        [Parameter( Mandatory = $true, ParameterSetName = "Object")]
        [ValidateNotNullOrEmpty()]
        [HashTable]
        $NameValuePairs,

        # String array of one or more sections to limit the changes to, separated by a comma.
        # Surrounding section names with square brackets is not necessary but is supported.
        # Ini keys that do not have a defined section can be modified by specifying '_' (underscore) for the section.
        [Parameter( ParameterSetName = "File" )]
        [Parameter( ParameterSetName = "Object" )]
        [ValidateNotNullOrEmpty()]
        [String[]]
        $Sections
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"

        # Update or add the name/value pairs to the section.
        Function Update-IniEntry {
            param ($content, $section)

            foreach ($pair in $NameValuePairs.GetEnumerator()) {
                if (!($content[$section])) {
                    Write-Verbose ("$($MyInvocation.MyCommand.Name):: '{0}' section does not exist, creating it." -f $section)
                    $content[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                }

                Write-Verbose ("$($MyInvocation.MyCommand.Name):: Setting '{0}' key in section {1} to '{2}'." -f $pair.key, $section, $pair.value)
                $content[$section][$pair.key] = $pair.value
            }
        }
    }
    # Update the specified keys in the list, either in the specified section or in all sections.
    Process {
        # Get the ini from either a file or object passed in.
        if ($PSCmdlet.ParameterSetName -eq 'File') { $content = Get-IniContent $FilePath }
        if ($PSCmdlet.ParameterSetName -eq 'Object') { $content = $InputObject }

        # Specific section(s) were requested.
        if ($Sections) {
            foreach ($section in $Sections) {
                # Get rid of whitespace and section brackets.
                $section = $section.Trim() -replace '[][]', ''

                Write-Debug ("Processing '{0}' section." -f $section)

                Update-IniEntry $content $section
            }
        }
        else {
            # No section supplied, go through the entire ini since changes apply to all sections.
            foreach ($item in $content.GetEnumerator()) {
                $section = $item.key

                Write-Debug ("Processing '{0}' section." -f $section)

                Update-IniEntry $content $section
            }
        }
        Write-Output $content
    }
    End {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}


function ShowToast 
    {
        ##############################
        #Author
        #@MattiasFors
        #https://deploywindows.com
        #https://github.com/DeployWindowsCom/DeployWindows-Scripts
        ##############################
        #
        #Ex:
        #ShowToast -ToastTitle "This is a Test" -ToastText "We toast are burned !!!" -ToastDuration Short #-Image "C:\temp\ekud.jpg"
        ##############################
        
        param(
                [parameter(Mandatory=$true,Position=2)]
                [string] $ToastTitle,
                [parameter(Mandatory=$true,Position=3)]
                [string] $ToastText,
                [parameter(Position=1)]
                [string] $Image = $null,
                [parameter()]
                [ValidateSet('long','short')]
                [string] $ToastDuration = "short"
            )


        Log-ScriptEvent -value "Preparing Toast notification message"
        # Toast overview: https://msdn.microsoft.com/en-us/library/windows/apps/hh779727.aspx
        # Toasts templates: https://msdn.microsoft.com/en-us/library/windows/apps/hh761494.aspx
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null

        # Define Toast template, w/wo image
        $ToastTemplate = [Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText02
        if ($Image.Length -le 0) { $ToastTemplate = [Windows.UI.Notifications.ToastTemplateType]::ToastText02 }

        # Download or define a local image. Toast images must have dimensions =< 1024x1024 size =< 200 KB
        if ($Image -match "http*") 
            {
                [System.Reflection.Assembly]::LoadWithPartialName("System.web") | Out-Null
                $Image = [System.Web.HttpUtility]::UrlEncode($Image)
                $imglocal = "$($env:TEMP)\ToastImage.png"
                Start-BitsTransfer -Destination $imglocal -Source $([System.Web.HttpUtility]::UrlDecode($Image)) -ErrorAction Continue
            } 
        else 
            { $imglocal = $Image }

        # Define the toast template and create variable for XML manipulation
        # Customize the toast title, text, image and duration
        $toastXml = [xml] $([Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(`
        $ToastTemplate)).GetXml()
        $toastXml.GetElementsByTagName("text")[0].AppendChild($toastXml.CreateTextNode($ToastTitle)) | Out-Null
        $toastXml.GetElementsByTagName("text")[1].AppendChild($toastXml.CreateTextNode($ToastText)) | Out-Null
        if ($Image.Length -ge 1) { $toastXml.GetElementsByTagName("image")[0].SetAttribute("src", $imglocal) }
        $toastXml.toast.SetAttribute("duration", $ToastDuration)

        # Convert back to WinRT type
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument; $xml.LoadXml($toastXml.OuterXml);
        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)

        # Get an unique AppId from start, and enable notification in registry
        $AppID = ((Get-StartApps -Name 'Windows Powershell') | Select -First 1).AppId
        New-Item "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\$AppID" -Force | Out-Null
        Set-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\$AppID" -Name "ShowInActionCenter" -Type Dword -Value "1" -Force | Out-Null
        # Create and show the toast, dont forget AppId
        Log-ScriptEvent -value "Sending notification to Windows"
        Try
            {
                [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppID).Show($Toast)
                Log-ScriptEvent -value "Notification sent !!!"
            }
        Catch
            {Log-ScriptEvent -value "[Error] Unable to send toast notification: $($_.Exception.Message)"}


    }


function New-ManagedCredential()
    {

        ###############################################################################################################
        # Language     :  PowerShell 4.0
        # Filename     :  ManagedCredentials.psm1
        # Autor        :  BornToBeRoot (https://github.com/BornToBeRoot)
        # Description  :  Secure encryption of credentials as SecureString
        # Repository   :  https://github.com/BornToBeRoot/PowerShell_ManagedCredential
        ###############################################################################################################

        <#
            .SYNOPSIS
            Secure encryption of credentials as SecureString
    
            .DESCRIPTION
            Secure encryption of credentials as SecureString, which can be saved as an xml-file or variable.

            If user "A" encrypt the credentials on computer "A", user "B" cannot decrypt the credentials on 
            computer "A" and also user "A" cannot decrypt the credentials on Computer "B".
        
            .EXAMPLE
            $EncryptedCredential = New-ManagedCredential

            $EncryptedCredential
        
            UsernameAsSecureString : c04fc297eb01000000edade3a984d5ca...
            PasswordAsSecureString : 984d5ca4aa6c39de63b9627730000c22...

            .EXAMPLE
            New-ManagedCredential -OutFile E:\Temp\EncryptedCredentials.xml

            .EXAMPLE
            New-ManagedCredential -OutFile E:\Temp\EncryptedCredentials.xml -overwrite

            .EXAMPLE
            New-ManagedCredential -UserName $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -Message "Please fill password for current logged-on user"

	        .LINK
	        https://github.com/BornToBeRoot/PowerShell_ManagedCredential/blob/master/Documentation/New-ManagedCredential.README.md
        #>

        [CmdletBinding(SupportsShouldProcess=$true)]
        Param(
                [Parameter(
                    Position=0,
                    HelpMessage='Path to the xml-file where the encrypted credentials will be saved')]
                [String]$OutFile,


                [Parameter(
		            Position=1,
		            HelpMessage='Credentials which are encrypted')]
                [System.Management.Automation.PSCredential]
                [System.Management.Automation.CredentialAttribute()]
                $Credential,

                [Parameter(
		            Position=2,
		            HelpMessage='Username that will be pre-populated')]
                [String]$UserName,

                [Parameter(
		            Position=3,
		            HelpMessage='Message displayed in the credential window')]
                [String]$Message
            )

        Begin
            {

	        }

	    Process
            {
		        if($null -eq $Credential)
                    {
                        try
                        {
                            If ([String]::IsNullOrWhiteSpace($UserName))
                                {$Credential = Get-Credential $null}
                            Else
                                {
                                    If ([String]::IsNullOrWhiteSpace($message)){$Message = "Please Enter valid credentials"}    
                                    $Credential = Get-Credential -UserName $UserName -Message $Message
                                }     
                        } 
			            catch
                        {
                            throw
                        }      
                    }
        
                $EncryptedUsername =  $Credential.UserName | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
                $EncryptedPassword =  $Credential.Password | ConvertFrom-SecureString

                $EncryptedCredentials = [pscustomobject] @{
                    UsernameAsSecureString = $EncryptedUsername
                    PasswordAsSecureString = $EncryptedPassword
                }
         
                if(-not([String]::IsNullOrEmpty($OutFile)))
                    {               
                        if(-not([System.IO.Path]::IsPathRooted($OutFile))) 
                            {$FilePath = Join-Path -Path $PSScriptRoot -ChildPath $OutFile.Replace(".\","")}
                        else
                            {$FilePath = $OutFile}
        
                        if(-not($FilePath.ToLower().EndsWith(".xml"))){$FilePath += ".xml"}

                        if($PSCmdlet.ShouldProcess($FilePath)){$EncryptedCredentials | Export-Clixml -Path $FilePath -Force}
                    }
                else 
                    {$EncryptedCredentials}
	        }

	    End
            {

	        }
    }


function Get-ManagedCredential()
    {

        ###############################################################################################################
        # Language     :  PowerShell 4.0
        # Filename     :  ManagedCredentials.psm1
        # Autor        :  BornToBeRoot (https://github.com/BornToBeRoot)
        # Description  :  Secure decryption of encrypted credentials
        # Repository   :  https://github.com/BornToBeRoot/PowerShell_ManagedCredential
        ###############################################################################################################

        <#
            .SYNOPSIS
	        Secure decryption of encrypted credentials

            .DESCRIPTION
   	        Secure decryption of encrypted credentials, which have been stored in an xml-file or variable.

            If user "A" encrypt the credentials on computer "A", user "B" cannot decrypt the credentials on 
            computer "A" and also user "A" cannot decrypt the credentials on Computer "B".

            .EXAMPLE
            Get-ManagedCredential -EncryptedCredential $EncryptedCredential

	        UserName                Password
	        --------                --------
	        Admin                   System.Security.SecureString

            .EXAMPLE
            Get-ManagedCredential -EncryptedCredential $EncryptedCredential AsPlainText

	        UserName                Password
	        --------                --------
	        Admin                   PowerShell

    
            .EXAMPLE
   	        Get-ManagerdCredential -FilePath E:\Temp\EncryptedCredentials.xml -AsPlainText

	        Username                Password
	        --------                --------
	        Admin                   PowerShell	
	        
            .LINK
            https://github.com/BornToBeRoot/PowerShell_ManagedCredential/blob/master/Documentation/Get-ManagedCredential.README.md
        #>

	    [CmdletBinding(DefaultParameterSetName='File')]
	    Param(
		    [Parameter(
			    ParameterSetName='File',
			    Position=0,
			    Mandatory=$true,
			    HelpMessage='Path to the xml-file where the encrypted credentials are saved')]
		    [ValidateScript({
			    if(Test-Path -Path $_ -PathType Leaf)
			    {
				    return $true
			    }
			    else 
			    {
				    throw "FilePath ($_) does not exist!"
			    }
		    })]
		    [String]$FilePath,
		
		    [Parameter(
			        ParameterSetName='Variable',
			        Position=0,
			        Mandatory=$true,
			        ValueFromPipeline=$true,
			        HelpMessage='Encrypted credential')]
		        [pscustomobject]$EncryptedCredential,
	
		        [Parameter(
			        Position=2,
			        HelpMessage='Return password as plain text')]
		        [Switch]$AsPlainText
	        )

	    Begin
        {

	    }

	    Process
	    {
		    if($PSCmdlet.ParameterSetName -eq 'File')
		        {			
			        try 
                    {$EncryptedCredential = Import-Clixml -Path $FilePath -ErrorAction Stop}
			        catch 
                    {throw}		
		        }
		
		    if($null -eq $EncryptedCredential){throw 'Nothing to decrypt. Try "Get-Help" for more details'}
		
		    try
            { 
                $SecureString_Username = $EncryptedCredential.UsernameAsSecureString | ConvertTo-SecureString -ErrorAction Stop
			    $SecureString_Password = $EncryptedCredential.PasswordAsSecureString | ConvertTo-SecureString -ErrorAction Stop
			

			    $BSTR_Username = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString_Username)
			    $Username = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR_Username) 

 
			    If ($AsPlainText) 
			        {
				        $BSTR_Password = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString_Password)
				        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR_Password)

				        [pscustomobject] @{
					        Username = $Username
					        Password = $Password
				        }
			        }
			    else
			        {New-Object System.Management.Automation.PSCredential($Username, $SecureString_Password)}
		    }
		    catch 
            {
			    throw	
		    }
	    }

	    End
        {

	    }
    }


	
Function Check-RightToWrite
	{
        param ([Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [System.String] $Path)		
		
		If (Test-Path $Path)
			{
				Try 
					{ 
						Log-ScriptEvent -value "#################Checking Access Right#################"
						Log-ScriptEvent -value "logged on as user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
						Log-ScriptEvent -value "Path to check: $Path"
						
						$FileStream = [System.IO.File]::OpenWrite($Path).close()
						Log-ScriptEvent -value "Write access to $Path checked successfully !!!" 
					}
		 		
				Catch [System.UnauthorizedAccessException]
					{Log-ScriptEvent -value "[ERROR] Unable to write to output file $outputfile" -Severity 3}
				Finally
					{
						Try
							{
								$FileStream = [System.IO.File]::Open($Path,'Open','Write')
								Log-ScriptEvent -value "File $Path is not locked"
							}
						Catch{Log-ScriptEvent -value "[ERROR] File $outputfile is locked !!!" -Severity 3}
						Finally
							{
								$FileStream.Close()
								$FileStream.Dispose()
							}
						
						Log-ScriptEvent -value "Effective permissions on target :" 
						ForEach ($Item in (Get-Acl $Path).access)
							{
								Log-ScriptEvent -value " "
							  	Log-ScriptEvent -value "Account: $($Item.IdentityReference)"
							  	Log-ScriptEvent -value "Permission: $($Item.AccessControlType)"
							  	Log-ScriptEvent -value "FileSystemRights: $($Item.FileSystemRights)"
							}	
					}
			}
		Else
			{Log-ScriptEvent -value "[ERROR]File/Folder $outputfile does not exists, aborting !!!!"  -Severity 3}
	}	


function Resolve-Error 
	{
	
		<# 
		 
		.SYNOPSIS 
		Recurses an error record or exception object to flatten nested objects. 
		 
		.DESCRIPTION 
		Loops through information caught in catch blocks; from an ErrorRecord (and its InvocationInfo), to Exception, and InnerException. 
		 
		.PARAMETER ErrorRecord 
		An error record or exception. By default the last error is used. 
		 
		.PARAMETER AsString 
		Return an array of strings for printable output. By default we return an array of objects. 
		 
		.PARAMETER Reverse 
		Returns items from outermost to innermost. By default we return items innermost to outermost. 
		 
		.INPUTS 
		By default the last error; otherwise any error record or exception can be passed in by pipeline or first parameter. 
		 
		.OUTPUTS 
		An array of objects, or an array of strings. 
		 
		.EXAMPLE 
		Resolve-Error 
		 
		Returns an array of nested objects describing the last error. 
		 
		.EXAMPLE 
		$_ | Resolve-Error

		Returns an array of strings describing the error in $_. 
		 
		#>	
	
	
		
		
	    [CmdletBinding()]
	    param (
	        [Parameter(ValueFromPipeline = $true)]
	        $ErrorRecord = $null,
	        [switch] $AsString = $True,
            [switch] $InLogFile = $true,
	        [switch] $Reverse
	    )

	    if (!$ErrorRecord) 
			{
		        # This is a bit iffy, if it's a nested module it needs $_ as $Error will not be populated yet.
		        # If it's not a nested module then it needs a Get-Variable -Scope 2 
		        $ErrorRecord = (Get-Variable -Name Error -Scope 2).Value | Select-Object -First 1
		        <# 
		        if ($Error.Count -gt 0) { 
		            $ErrorRecord = $Error[0] 
		        } else { 
		            $ErrorRecord = $_ 
		        } 
		        #>
	    	}	
	    
		$records = @()
	    
	    if ($ErrorRecord.psobject.Properties["InnerException"] -and $ErrorRecord.InnerException) {$records += Resolve-Error $ErrorRecord.InnerException}
	    if ($ErrorRecord.psobject.Properties["Exception"] -and $ErrorRecord.Exception) {$records += Resolve-Error $ErrorRecord.Exception}
	    if ($ErrorRecord.psobject.Properties["InvocationInfo"] -and $ErrorRecord.InvocationInfo) {$records += Resolve-Error $ErrorRecord.InvocationInfo}
	    $records += $ErrorRecord

	    if ($Reverse) {$records = [Array]::Reverse($records) ; $AsString = $False}
	    if (!$AsString)
            {$records} 
		else 
			{
		        $string = @()
		        $first = $true

		        $records | ForEach-Object {
			            if ($first) 
							{
			                	$string += "=" * 40
			                	$first = $false
			            	} 
						else 
                            {$string += "*" * 5}
			            $string += $_ | Select-Object * | Out-String
			        }

	    		$string += ""

		        $stack = Get-PSCallStack
		        for ($i = $stack.Count - 1; $i -ge 1; $i--) 
					{ 
			            $string += "-" * 5
			            $string += "Depth: $i"
			            $string += "Function: $($stack[$i].FunctionName)"
			            # In some highly threaded contexts this doesn't appear?
			        	if ($stack[$i].PSObject.Properties["Arguments"]) 
							{
			               		$string += "Arguments: $($stack[$i].Arguments)"
			            	}
		        		$string += "Line: $($stack[$i].ScriptLineNumber)"
		            	$string += "Command: $($stack[$i].Position.Text)"
	        		}
		        $string += ""
		        $string += ""
		        $string += "=" * 40
                $string -join [System.Environment]::NewLine

                If ($InLogFile)
                    {
                        $Lines = @()
                        ForEach ($Item in $String)
                            {
                                $Lines = $Item -split "`r`n"
                                $Lines|Foreach {If (-not([String]::IsNullOrWhiteSpace($_))){Log-ScriptEvent $_ -Severity 3}}
                            }
                    }
        	}
	}	


Function Add-XmlFragment
	{
		<#
		.SYNOPSIS
		    Adds an XML fragment to an XmlNode
		    
		.DESCRIPTION
		    Adds an XML fragment to an XmlNode
		    
		.NOTES 
		    Author     : Thomas Freudenberg - info@thomasfreudenberg.com
		    
		.EXAMPLE
			$OSGroupxmlFile = Get-Content -path "$DSPath\Control\OperatingSystemGroups.xml"
			[xml]$XMLNewGroup =@"
			<group enable="True" guid="{$([guid]::NewGuid())}"><Name>$FFUGroup</Name> ... </group>
			"@
			Add-XmlFragment -XmlElement $OSGroupxmlFile.groups  -XmlNodeToAdd [xml]$XMLNewGroup
			$OSGroupxmlFile.Save("$DSPath\Control\OperatingSystemGroups.xml")
		    
		    Adds the connection string "MyConString" to the web.config file.
		#>

		    Param(        
		        	[Parameter(Mandatory=$true)][System.Xml.XmlNode] $XmlElement,
		        	[Parameter(Mandatory=$true)][string] $XmlNodeToAdd
				)
				
		    $xml = $xmlElement.OwnerDocument.ImportNode(([xml]$XmlNodeToAdd).DocumentElement, $true)
		    [void]$xmlElement.AppendChild($xml)
	}


function convertFileSize {
    param(
        $bytes
    )

    if ($bytes -lt 1MB) {
        return "$([Math]::Round($bytes / 1KB, 2)) KB"
    }
    elseif ($bytes -lt 1GB) {
        return "$([Math]::Round($bytes / 1MB, 2)) MB"
    }
    elseif ($bytes -lt 1TB) {
        return "$([Math]::Round($bytes / 1GB, 2)) GB"
    }
}


Function  Invoke-HTTPDownload
    {  
         <#PSScriptInfo

        .VERSION 19.05.24
        .GUID d1de0de3-d0f2-4094-a2f2-f75cd8d0fab3
        .AUTHOR Tim Small
        .COMPANYNAME
        .COPYRIGHT 2019
        .TAGS web download webclient
        .LICENSEURI
        .PROJECTURI https://github.com/Smalls1652/Get-RandomPowerShellScripts/blob/master/Actions/Networking/Invoke-DownloadFile.ps1
        .ICONURI
        .EXTERNALMODULEDEPENDENCIES
        .REQUIREDSCRIPTS
        .EXTERNALSCRIPTDEPENDENCIES
        .RELEASENOTES
        Fixed current dir paths
        .PRIVATEDATA

        .DESCRIPTION
         Download files from the internet through PowerShell.

        #>

        param(
            [Parameter(Mandatory = $true)][string]$Url,
            [string]$Path
        )

        $Status = $False

        #Load in the WebClient object.
        try 
            {$Downloader = New-Object -TypeName System.Net.WebClient}
        catch [Exception] 
            {
                Log-ScriptEvent "[Error] Unable to register Webclient Com object, aborting !!!" -Severity 3
                Exit
            }

        #Creating a temporary file if not in WinPE.
        If (-not $OSD_Env.IsWinPE)
            {
                $TmpFile = New-TemporaryFile
                Log-ScriptEvent "Download TmpFile set to ""$($TmpFile)""."
            }


        try 
            {
                #Start the download by using WebClient.DownloadFileTaskAsync, since this lets us show progress on screen.
                Log-ScriptEvent "Starting HTTP download..."
                If (-not $OSD_Env.IsWinPE)
                    {$FileDownload = $Downloader.DownloadFileTaskAsync($Url, $TmpFile)}
                Else
                    {$FileDownload = $Downloader.DownloadFileTaskAsync($Url, $Path)}

                #Register the event from WebClient.DownloadProgressChanged to monitor download progress.
                Register-ObjectEvent -InputObject $Downloader -EventName DownloadProgressChanged -SourceIdentifier WebClient.DownloadProgressChanged | Out-Null

                #Wait two seconds for the registration to fully complete
                Start-Sleep -Seconds 4

                if ($FileDownload.IsFaulted) 
                    {
                        Log-ScriptEvent "An error occurred. Unable to initate download." -Severity 2
                        break
                    }

                #While the download is showing as not complete, we keep looping to get event data.
                while (!($FileDownload.IsCompleted)) 
                    {

                        if ($FileDownload.IsFaulted) 
                            {
                                Log-ScriptEvent "An error occurred. Unable to download." -Severity 2
                                Write-Error $FileDownload.GetAwaiter().GetResult()
                                break
                            }

                        $EventData = Get-Event -SourceIdentifier WebClient.DownloadProgressChanged | Select-Object -ExpandProperty "SourceEventArgs" -Last 1

                        $ReceivedData = ($EventData | Select-Object -ExpandProperty "BytesReceived")
                        $TotalToReceive = ($EventData | Select-Object -ExpandProperty "TotalBytesToReceive")
                        $TotalPercent = $EventData | Select-Object -ExpandProperty "ProgressPercentage"

                        $Message = "Downloading file $(convertFileSize -bytes $ReceivedData) / $(convertFileSize -bytes $TotalToReceive)"
                        Show-TSActionProgress -Message $Message -Step $ReceivedData -MaxStep $TotalToReceive
                        Log-ScriptEvent $Message
                    }
            }
        catch [Exception] 
            {$_ | Resolve-Error}
        finally 
            {
                
                if (($FileDownload.IsCompleted) -and !($FileDownload.IsFaulted)) 
                    {
                        #Cleanup tasks
                        Log-ScriptEvent "Downloading file completed !"
                        Show-TSActionProgress -Message "Downloading file completed !" -Step 100 -MaxStep 100
                        Unregister-Event -SourceIdentifier WebClient.DownloadProgressChanged -force -ErrorAction SilentlyContinue

                        #If the download was finished without termination, then we move the file.
                        If (-not $OSD_Env.IsWinPE)
                            {
                                Log-ScriptEvent "Moving the downloaded file to ""$($Path)""."
                                Move-Item -Path $TmpFile -Destination $Path -Force
                            }
                        $Status = $True
                    }
                else 
                    {
                        #If the download was terminated, we remove the file.
                        Unregister-Event -SourceIdentifier WebClient.DownloadProgressChanged -force -ErrorAction SilentlyContinue
                        Log-ScriptEvent "Cancelling the download and removing the tmp file (If any)."
                        $Downloader.CancelAsync()
                        If (-not $OSD_Env.IsWinPE){Remove-Item -Path $TmpFile -Force} Else {Remove-Item -Path $Path -Force}
                        $Status = $False
                    }
            } 
            
        Remove-Event -SourceIdentifier "WebClient.DownloadProgressChanged" -ErrorAction SilentlyContinue
        $Downloader.Dispose()      
        Return $Status
}  



Function Invoke-BITSDownload
    {
        param (
            [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [System.String] $url,
            [Parameter(Mandatory)] [ValidateNotNullOrEmpty()] [System.String] $Path
            )
        
        $Job = Start-BitsTransfer -Source $url -Destination $Path -Asynchronous
        Log-ScriptEvent -value "Starting BITS download of $url"

        While( ($Job.JobState.ToString() -eq 'Transferring') -or ($Job.JobState.ToString() -eq 'Connecting') )
            {
                $ReceivedData = $Job.BytesTransferred
                $TotalToReceive = $Job.BytesTotal
                $pct = [int](($ReceivedData*100) / $TotalToReceive)
                $Message = "Downloading file $(convertFileSize -bytes $ReceivedData) / $(convertFileSize -bytes $TotalToReceive)"
                Show-TSActionProgress -Message $Message -Step $ReceivedData -MaxStep $TotalToReceive
                Log-ScriptEvent $Message
                Start-Sleep -Seconds 1
            }

        Switch($Job.JobState)
            {
                "Transferred" 
                    {
                        Complete-BitsTransfer -BitsJob $Job
                        Show-TSActionProgress -Message "Downloading file completed" -Step 100 -MaxStep 100
                        Log-ScriptEvent -value "Download job Finished... 100% complete"
                        Return $true
                    }
                "Error" 
                    {
                        #$Job | Format-List 
                        Log-ScriptEvent "An error occurred. Unable to complete download." -Severity 2
                        Return $false
                    } # List the errors.

            }

    } 


Function Download-File
    {
        param(
                [Parameter(Mandatory = $true)][string]$Url,
                [string]$Path
            )

        Log-ScriptEvent "Download URL set to ""$($Url)""."

        switch ($Path) 
            {
                { ([regex]::Match($PSItem, "(?:(?:.+[\/\\]){2,}(?:(?'fileName'.+)|)|.+)").Groups | Where-Object -Property "Name" -eq "fileName" | Select-Object -ExpandProperty "Success") -eq $false } 
                    {
                        Log-ScriptEvent "Path parameter set, but no filename. Parsing Url for filename."
                        if (!([regex]::Match($Path, "^.+[\/\\]$") | Select-Object -ExpandProperty "Success")) 
                            {
                                if ($Path.ToCharArray() -contains "/") 
                                    {$Path = "$($Path)/"}
                                elseif ($Path.ToCharArray() -contains "\")
                                    {$Path = "$($Path)\"}
                            }
                        $URLParser = $Url | Select-String -Pattern ".*\:\/\/.*\/(.*\.{1}\w*).*" -List

                        $Path = "$($Path)$($URLParser.Matches.Groups[1].Value)"
                        Break
                    }

                { ([regex]::Match($PSItem, "(?:(?:.+[\/\\]){2,}(?:(?'fileName'.+)|)|.+)").Groups | Where-Object -Property "Name" -eq "fileName" | Select-Object -ExpandProperty "Success") -eq $true } 
                    {
                        Log-ScriptEvent "Path parameter already set, proceeding to download stage!"
                        Break
                    }

                Default 
                    {
                        Log-ScriptEvent "Path parameter not set, parsing Url for filename."
                        $URLParser = $Url | Select-String -Pattern ".*\:\/\/.*\/(.*\.{1}\w*).*" -List

                        $Path = "./$($URLParser.Matches.Groups[1].Value)"
                    }
            }

        Log-ScriptEvent "Save Path set to ""$($Path)""."

        $iRet = $false ; $retry = 1

        ## Try downloading with HTTP, Wait 60s if it fails and retry 2 times
        While ($iRet -ne $true -and $retry -le 3 )        
            {
                Log-ScriptEvent "Downloading content with HTTP, attempt $retry"
                $iRet = Invoke-HTTPDownload -Url $Url -Path $path
                If ($iRet -eq $true){Break}
                Start-Sleep -Seconds 60
                $retry += 1
            } 
        
        $retry = 1

        ## Try downloading with BITS, Wait 60s if it fails and retry 2 times
        While ($iRet -ne $true -and $retry -le 3 )        
            {
                Log-ScriptEvent "Downloading content with BITS, attempt $retry"
                $iRet = Invoke-HTTPDownload -Url $Url -Path $path
                If ($iRet -eq $true){Break}
                Start-Sleep -Seconds 60
                $retry += 1
            } 

        If ($iret -eq $true)
            {
                Log-ScriptEvent "File Downloaded successfully !!"
                Log-ScriptEvent "File path : $Path"
                Return $Path
            }
        Else
            {
                Log-ScriptEvent "[ERROR] Aborting Download of $Url , after 6 retry !!"
                Return $False
            }
    }



##################################################################################