param([parameter(Mandatory = $True)][string]$DeployShare, [switch]$RestoreBeforeProceed = $True )

#########################

# MDT'O'Matic
# add Parachute and other more or less useful functionality to MDT 

# By Diagg/OSDC

# Version 2.0 - Release date: 09.26.2021 - 7zip is replaced by Wimlib
# 										 - Prevent from running on Powershell 7.
#                                        - ISO retrieval is now done during MDT ZTI-Gather.wsf processing.

#########################


#Requires -Version 4
#Requires -RunAsAdministrator 

##== Debug
$ErrorActionPreference = "stop"
#$ErrorActionPreference = "Continue"
$Error.Clear()

##== Global Variables
$Script:CurrentScriptName = $MyInvocation.MyCommand.Name
$Script:CurrentScriptFullName = $MyInvocation.MyCommand.Path
$Script:CurrentScriptPath = split-path $MyInvocation.MyCommand.Path

##== Check Powershell Version
If ($PSVersionTable.PSVersion.Major -ge 7)
	{
        (New-Object -ComObject wscript.shell).popup("   [ERROR] This script does not support Powershell 7.`r`r   Aborting !!`r",0,"Warning",0x0 +16)
        Write-Error "[ERROR] This script does not support Powershell 7, Aborting !!"
        Exit		
	}



##== Check Deployment share
If (-not(Test-path "$DeployShare\Control\OperatingSystems.xml") -and (-not(Test-path "$DeployShare\Scripts\LTIApply.wsf")))
    {
        (New-Object -ComObject wscript.shell).popup("   [ERROR] Not a valid Deployment share.`r`r   Aborting !!`r",0,"Warning",0x0 +16)
        Write-Error "[ERROR] Not a valid Deployment share, Aborting !!"
        Exit
    }

If (-not (Test-Path "C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1"))
    {
        (New-Object -ComObject wscript.shell).popup("   [ERROR] MDT is not installed on this computer.`r`r   Aborting !!`r",0,"Warning",0x0 +16)
        Write-Error "[ERROR] MDT is not installed on this computer, dependencies are missing, Aborting !!"
        Exit
    }


##== Importing MDT Module
If ([string]::IsNullOrWhiteSpace($(Get-Module -Name "MicrosoftDeploymentToolkit")))
    {
        Import-Module "C:\Program Files\Microsoft Deployment Toolkit\Bin\MicrosoftDeploymentToolkit.psd1"
    }

##== Get Associated deployment share
Restore-MDTPersistentDrive
ForEach ($Item in $(Get-MDTPersistentDrive))
    {
        If ($item.path.ToUpper() -eq $DeployShare.ToUpper())
            {
                $DeployDrive ="$($item.Name):"
                Break
            }
    }

If ([string]::IsNullOrWhiteSpace($DeployDrive))
    {
        New-PSDrive -Name "DSManager" -PSProvider MDTProvider -Root $DeployShare
        $DeployDrive = "DSManager:"
    }


#########################
##### Archived Data #####
#########################

$ArchivedContent="##EMBED-HERE##"

##== Decoding
$Content = [System.Convert]::FromBase64String($ArchivedContent)
Set-Content -Path "$env:temp\Embedded.zip" -Value $Content -Encoding Byte

##== unzipping
$Hangar18 = "$Script:CurrentScriptPath\Extracted"
If (-not(Test-path $Hangar18)){New-Item -Path $Hangar18 -ItemType Directory -Force|out-null}
Expand-Archive -LiteralPath "$env:temp\Embedded.zip" -DestinationPath $Hangar18 -Force
Remove-Item "$env:temp\Embedded.zip"

##== Includes
."$Hangar18\DiaggFunctions.ps1"
Init-Function

##== Variables
$RestoreSource ='C:\Program Files\Microsoft Deployment Toolkit\Templates\Distribution'
$UpdateDeploymentShare = $false
$Backup
$BootImageFiles = @("ZTIUtility.vbs","ZTIGather.wsf","ZTIGather.xml","Wizard.css","LiteTouch.wsf","ltiCleanup.wsf","Bootstrap.ini")


#########################
##### Populate MDT  #####
#########################

##== Create OS folder
If(-not (test-path "$DeployDrive\Operating Systems\Online OS")){New-Item -path "$DeployDrive\Operating Systems" -enable “True” -Name “Online OS” -comments "Images downloaded from the cloud" -ItemType “folder”|Out-Null }

##==Import OS
If (-not (Test-Path "$DeployDrive\Operating Systems\Online OS\Windows 10 Pro in Windows 10 Latest Online install.wim")){Import-MdtOperatingSystem -path "$DeployDrive\Operating Systems\Online OS" -SourceFile "$Hangar18\install.wim" -DestinationFolder "Windows 10 Latest Online" | Out-Null}   

#################################
##### Data to Edit #####
#################################

$DataList = New-Object System.Collections.ArrayList


#################
#### LTIApply.wsf
#################

$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value 'ec1691db-5647-42fe-bdc7-a6ba128d2017'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:3
This:allow OS Index gathered from xml to be replaced by an MDT variable
GUID:$($Item.GUID)
"@

$StringInsert = @"
		If oEnvironment.Exists("MountedIsoWIMIndex") then
			sImageIndex = oEnvironment.Item("MountedIsoWIMIndex")
		Else
			sImageIndex = oUtility.SelectSingleNodeString(oOS,"ImageIndex")		
		End If	
		'//  MDT-O-Matic: $($Item.GUID)
"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\LTIApply.wsf"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'REPLACE_LINE'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 827
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $true
$DataList.Add($Item)|Out-Null

##########

$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value '76b5008a-bb85-42e5-ab0d-294ec9851a60'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:3
This:allow OS path gathered from xml to be replaced by an MDT variable
GUID:$($Item.GUID)
"@

$StringInsert = @"

		If oEnvironment.Exists("MountedIsoWIMPath") then
			GetSourcePath = oEnvironment.Item("MountedIsoWIMPath")
		End If	
		'//  MDT-O-Matic: $($Item.GUID)

"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\LTIApply.wsf"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'INSERT'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 439
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $false
$DataList.Add($Item)|Out-Null

###################
#### ZTIGather.wsf
###################
$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value '387dbe8c-c963-4089-93c7-aba5ca9aec85'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:4
This:Gather about Internet Function
GUID:$($Item.GUID)
"@

$StringInsert = @"


	'//  MDT-O-Matic: $($Item.GUID)
	'//---------------------------------------------------------------------------
	'//  Function:	GetInternetInfo()
	'//  Purpose:	Gather if we can go to the internet or not.
	'//---------------------------------------------------------------------------
	Function GetInternetInfo
	
		Dim oHTTP, oHttpReply

		oLogging.CreateEntry "Getting Internet info", LogTypeInfo 

		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		oHTTP.Open "GET", "https://www.microsoft.com/en-us/software-download/windows10ISO", FALSE
		On Error Resume Next
		oHTTP.Send
		oHttpReply = oHTTP.statusText
		If Err.Number <> 0 Then 
			oEnvironment.Item("IsInternetConnected") = oUtility.ConvertBooleanToString(false)
			oLogging.CreateEntry "Internet not accessible - unable to reach a web page!", LogTypeInfo
		Elseif oHttpReply = "OK" Then 
			oLogging.CreateEntry "Internet is " & oHttpReply & " - connexion/reply to web page was done succefully!", LogTypeInfo
			oEnvironment.Item("IsInternetConnected") = oUtility.ConvertBooleanToString(true)
		End If	
		On Error Goto 0
		
		GetInternetInfo = SUCCESS
		oLogging.CreateEntry "Finished getting Internet info", LogTypeInfo
		
	End Function		


	'//---------------------------------------------------------------------------
	'//  Function:	GetOnlineOSInfo()
	'//  Purpose:	Gather OS version we can download from internet.
	'//---------------------------------------------------------------------------
	Function GetOnlineOSInfo
	
		Dim oHTTP, oHttpReply, oWinVersion, oWinVersionName, oRegex, oMatches, oUrl, oLanguages, oLines ,i , oLanguageList , iRetVal

		oLogging.CreateEntry "Getting Online OS info", LogTypeInfo 
		
        Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		oHTTP.Open "GET", "https://www.microsoft.com/en-us/software-download/windows10ISO", FALSE
		On Error Resume Next
		oHTTP.Send
		oHttpReply = oHTTP.statusText
		If Err.Number <> 0 Then 
			oLogging.CreateEntry "Internet not accessible - unable to reach a web page!", LogTypeInfo
			Exit Function
		End If		

		If not oEnvironment.Exists("LatestOnlineVersion") Then

			oWinVersion = oHTTP.responseText

			' Prepare Regex
			Set oRegex = New RegExp
			oRegex.Global  = False
			oRegex.IgnoreCase = True

			' Get Latests Online Windows 10 version
			'oRegex.Pattern = "<select id=""product-edition""([\s\S]*?)</select>"
			oRegex.Pattern = "<optgroup([\s\S]*?)</optgroup>"
			Set oMatches = oRegex.Execute(oWinVersion)
			oWinVersion = oMatches(0)
			oWinVersionName = oMatches(0)
			Set oMatches = Nothing
			
			' Get Friendly Name
			oRegex.Pattern = "<optgroup label=""([\s\S]*?)"""
			Set oMatches = oRegex.Execute(oWinVersionName)
			oWinVersionName = oMatches(0)
			oWinVersionName = Replace(oWinVersionName,"""","")
			oWinVersionName=Split(oWinVersionName,"=")(1)	
			Set oMatches = Nothing

			' Get Internal version Number
			oRegex.Pattern = "value=""([\s\S]*?)"""
			Set oMatches = oRegex.Execute(oWinVersion)
			oWinVersion = oMatches(0)
			oWinVersion = Replace(oWinVersion,"""","")
			oWinVersion=Split(oWinVersion,"=")(1)
			Set oMatches = Nothing			
			
			oEnvironment.Item("LatestOnlineVersion") = oWinVersion
			oEnvironment.Item("LatestOnlineVersionName") = oWinVersionName	

			' Query Available online languages
			oUrl = "https://www.microsoft.com/en-us/api/controls/contentinclude/html"
			oUrl = oUrl & "?pageId=a8f8f489-4c7f-463a-9ca6-5cff94d8d041"
			oUrl = oUrl & "&host=www.microsoft.com"
			oUrl = oUrl & "&segments=software-download,windows10ISO" 
			oUrl = oUrl & "&query=&action=getskuinformationbyproductedition"
			oUrl = oUrl & "&sessionId=" & left( CreateObject("Scriptlet.TypeLib").GUID, 38 )
			oUrl = oUrl & "&productEditionId=" & oEnvironment.Item("LatestOnlineVersion")
			oUrl = oUrl & "&sdVersion=2"
			oHttp.open "GET", oUrl, false
			oHttp.send
			oLanguages = oHttp.responseText

			' Get Language liste
			oRegex.Pattern = "<Select([\s\S]*?)</select>"
			Set oMatches = oRegex.Execute(oLanguages)
			oLanguages = oMatches(0)
			Set oMatches = Nothing
			
			oLanguages = Replace(oLanguages,"&quot;","")
			oLanguages = Replace(oLanguages,"<select id=""product-languages"">","")
			oLanguages = Replace(oLanguages,"</select>","")
			oLanguages = Replace(oLanguages,"<select id=""product-languages"">","")
			oLanguages = Replace(oLanguages,"<option value="""" selected=""selected"">Choose one</option>","")

			' Parse Languages
			oRegex.Pattern = "e:([\s\S]*?)}"			
			oLines=split(oLanguages,vbcrlf)

			For i=0 to ubound(olines)
				Set oMatches = oRegex.Execute(olines(i))
				If oMatches.count > 0 Then 
					oLanguages = Replace (oMatches(0),"e:","")
					oLanguages = Replace (oLanguages,"}","")
					oLanguageList  = oLanguageList & "," & oLanguages
				End If	
			next			

			oEnvironment.Item("RawOnlineLanguageList") = mid(oLanguageList,2)

		End If 	

		GetOnlineOSInfo = 0
		oLogging.CreateEntry "Finished online Os version info", LogTypeInfo
		
	End Function
	'//  MDT-O-Matic: $($Item.GUID)
	
"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\ZTIGather.wsf"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'INSERT'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 2196
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $true
$DataList.Add($Item)|Out-Null


########
$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value '84776031-487f-485b-bd1f-8a725aa211cd'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:5
This:Call Internet Function
GUID:$($Item.GUID)
"@

$StringInsert = @"

			'//  MDT-O-Matic: $($Item.GUID)	
			oLogging.ReportProgress "Getting local computer information: Internet connectivity", 50
			On Error Resume Next
			iRetVal = GetInternetInfo
			TestAndLog iRetVal ,"GetInternetInfo for Gather process"
			On Error Goto 0
	
			If Ucase(oEnvironment.Item("IsInternetConnected")) = "TRUE" Then
			
				oLogging.ReportProgress "Getting Online information: OS Availability", 60
				On Error Resume Next
				iRetVal = GetOnlineOSInfo
				TestAndLog iRetVal ,"GetOnlineOSInfo for Gather process"
				On Error Goto 0			
			
			End If
			'//  MDT-O-Matic: $($Item.GUID)
			
"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\ZTIGather.wsf"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'INSERT'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 146
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $false
$DataList.Add($Item)|Out-Null


#################
#### DeployWiz_Definition_ENU.xml
#################
$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value 'bbc248fb-f0b0-47e3-ace2-6c7194ccbe93'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:6
This:Add New language Wizard with available online ISO images
GUID:$($Item.GUID)
"@

$StringInsert = @"


	<!-- MDT-O-Matic: $($Item.GUID) -->
	<Pane id="SelectLanguageUI" reference="DeployWiz_LanguageUI.xml">
		<Condition><![CDATA[ UCase(Property("SkipLocaleSelection")) <> "YES" or UCase(Property("SkipTimeZone"))<>"YES" ]]> </Condition>
		<Condition><![CDATA[ Property("DeploymentType")<>"REPLACE" and Property("DeploymentType")<>"CUSTOM" and Property("DeploymentType") <> "StateRestore" and Property("DeploymentType")<> "UPGRADE" ]]> </Condition>
		<Condition><![CDATA[ FindTaskSequenceStep("//step[@name='Download Image']", "ZTIPowerShell.wsf" )<>True  ]]> </Condition>
	</Pane>
    <!-- MDT-O-Matic: $($Item.GUID) -->

			
"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\DeployWiz_Definition_ENU.xml"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'INSERT'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 90
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $True
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeChecked' -Value $false
$DataList.Add($Item)|Out-Null


########
$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value '35720ae6-2401-44b7-85ce-010623f1c48d'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:7
This:Add condition to the old language Wizard
GUID:$($Item.GUID)
"@

$StringInsert = @"

		<Condition><![CDATA[ FindTaskSequenceStep("//step[@name='Download Image']", "ZTIPowerShell.wsf" )<>True  ]]> </Condition>
		<!-- MDT-O-Matic: $($Item.GUID) -->
"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\DeployWiz_Definition_ENU.xml"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'INSERT'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 88
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $false
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeChecked' -Value $false
$DataList.Add($Item)|Out-Null



#################
#### Gather.xml
#################

$Item = New-Object PSObject
$Item|Add-Member -MemberType NoteProperty -Name 'GUID' -Value '2e82879f-a445-4ce3-bb0e-a19295b7e31d'

$Purpose = @"
General:Get Latest Windows ISO from Internet
Part:3
This:Add a new MDT variable called 
GUID:$($Item.GUID)
"@

$StringInsert = @"
	<property id="EditionToInstall" type="string" overwrite="true" description="Edition of Windows to Install In Cloud deployment scenario" />
	<!--  MDT-O-Matic: $($Item.GUID) -->
"@

$Item|Add-Member -MemberType NoteProperty -Name 'Purpose' -Value $Purpose
$Item|Add-Member -MemberType NoteProperty -Name 'File' -Value "\Scripts\ZTIGather.xml"
$Item|Add-Member -MemberType NoteProperty -Name 'Method' -Value 'INSERT'
$Item|Add-Member -MemberType NoteProperty -Name 'InsertLine' -Value 15
$Item|Add-Member -MemberType NoteProperty -Name 'StringInsert' -Value $StringInsert
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeRestored' -Value $true
$Item|Add-Member -MemberType NoteProperty -Name 'CanBeChecked' -Value $false
$DataList.Add($Item)|Out-Null

##########


############################# 
##== Process various Changes
#############################

Foreach ($Obj_item in $DataList)
    {
        $File = "$DeployShare$($Obj_item.File)"
        $RestoreFile = "$RestoreSource$($Obj_item.File)"
        $FriendlyName =  Split-Path $File -leaf
        
        ##== Backup/Restore if needed!
        If ($RestoreBeforeProceed -and $Obj_item.CanBeRestored)
            {
                Log-ScriptEvent "Restoring original file $($Obj_item.File) from $RestoreSource"
                Copy-Item -Path $RestoreFile -Destination $File -Force|out-Null
            }

        ##== Installing new files
        Log-ScriptEvent "Editing  File $FriendlyName at line $($Obj_item.InsertLine) !"  
        $Obj_item.Purpose.Split([Environment]::NewLine)|ForEach-Object {If (-Not([String]::IsNullOrWhiteSpace($_))){Log-ScriptEvent $_}}
        
              
        $content = Get-Content -Path $File

        If ($Obj_item.CanBeChecked -ne $false)
            {
                If (-not ($content| Where-Object { $_.Contains('6.3.8456.1000')}))
                    {
                        Log-ScriptEvent "File version unsupported for file $FriendlyName, version expected is 6.3.8456.2000, No modifiaction applied, Skipping to next file !" -Severity 2
                        Continue         
                    }
            }


        If (-not($content| Where-Object { $_.Contains($Obj_item.GUID)}))
            {
                 switch ($Obj_item.Method)
                    {
                        'REPLACE_STRING'
                            {
                                $newContent = $content -replace $Obj_item.StringtoReplace, $Obj_item.StringReplacement
                                $newContent | Set-Content -Path $File
                                Break
                            }

                        'REPLACE_LINE'
                            {
                                $content[$($Obj_item.InsertLine)-1] = $Obj_item.StringInsert
                                $Content | Set-Content -Path $File
                                Break
                            }
                
                        'INSERT'
                            {
                                $content[$($Obj_item.InsertLine)-1] += $Obj_item.StringInsert
                                $Content | Set-Content -Path $File
                                Break
                            }                    
                    }

                Log-ScriptEvent " + $FriendlyName file modified succefully !"
            }
        Else
            {Log-ScriptEvent " - $FriendlyName file already modified !"}

        If ($BootImageFiles -contains $FriendlyName){$UpdateDeploymentShare = $True}
    }


#################################
##### Data to Download #####
#################################

##== Downloading WimLib command line tool
$URL = "https://wimlib.net/downloads/wimlib-1.13.4-windows-x86_64-bin.zip"
$WimlibExtra = "$Hangar18\wimlib-1.13.4-windows-x86_64-bin.zip"

$webclient = New-Object System.Net.WebClient
$webclient.DownloadFile($url,$WimlibExtra)

##== Extract downloaded content

[System.IO.Compression.ZipFile]::ExtractToDirectory($WimlibExtra,"$Hangar18\Wimlib")
$WimLibExtra = "$Hangar18\Wimlib"

#################################
##### Data to Copy #####
#################################

##==Copy to Scripts folder
$Files = [PSCustomObject]@{ Source = "$Hangar18" ; Destination = "$DeployShare\Scripts" ; FileList = @("DeployWiz_LanguageUIOnline.vbs","DeployWiz_LanguageUIOnline.xml","DiaggFunctions.ps1","Download-Iso.ps1")}
ForEach ($item in $Files.FileList)
    {
        Copy-Item -Path "$($files.source)\$item" -Destination "$($files.Destination)\$item" -Force
        Log-ScriptEvent "Copying file $Item to $($files.Destination)" 
    }

##== Copy 7zip to Tools\x64 folder
Copy-Item -Path $WimLibExtra -Destination "$DeployShare\tools\x64" -Recurse -Force
Log-ScriptEvent "Copying folder $WimLibExtra to $DeployShare\tools\x64"

##== Copy TS to C:\Program Files\Microsoft Deployment Toolkit\Templates
Copy-Item -Path "$Hangar18\CloudClient.xml" -Destination "C:\Program Files\Microsoft Deployment Toolkit\Templates" -Force
Log-ScriptEvent "Copying file CloudClient.xml to C:\Program Files\Microsoft Deployment Toolkit\Templates"


###########################################
##== Update Deployment Share if needed ==##
###########################################

If ($UpdateDeploymentShare)
    {
        Log-ScriptEvent "Updating Boot Image, please Wait" 
        Update-MDTDeploymentShare -Path $DeployDrive -Verbose
    }