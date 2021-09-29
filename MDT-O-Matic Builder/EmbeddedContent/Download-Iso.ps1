#########################

## Windows 10 ISO Downloader 2.0
## Download latest available version

## Most of the ISO URL retrival code by Pete Batard <pete@akeo.ie>
## https://github.com/pbatard/Fido/blob/master/Fido.ps1 
## File download part by Tim Small 
## https://www.powershellgallery.com/packages/Invoke-DownloadFile/19.5.24/Content/Invoke-DownloadFile.ps1
## Put Together + MDT intégration by Diagg/OSDC
## V1 - Release 22/01/2021
## V2 - Release 03/08/2021
##      7Zip was replaced by WimLib

## Prefered Language possible value (depend on what Microsoft will make available): 
## Arabic, Brazilian Portuguese, Bulgarian, Chinese (Simplified), Chinese (Traditional), Croatian, Czech
## Danish, Dutch, English, English International, Estonian, Finnish, French, French Canadian, German, Greek
## Hebrew, Hungarian, Italian, Japanese, Korean, Latvian, Lithuanian, Norwegian, Polish, Portuguese, Romanian
## Russian, Serbian Latin, Slovak, Slovenian, Spanish, Spanish (Mexico), Swedish, Thai, Turkish, Ukrainian

## PreferedArchitecture possible value:
## x64, x86

#########################


#Requires -Version 4
##Requires -RunAsAdministrator 

##== Debug
$ErrorActionPreference = "stop"
#$ErrorActionPreference = "Continue"
$error.clear()

##== Global Variables
$Script:CurrentScriptName = $MyInvocation.MyCommand.Name
$Script:CurrentScriptFullName = $MyInvocation.MyCommand.Path
$Script:CurrentScriptPath = split-path $MyInvocation.MyCommand.Path

##== Includes
."$Script:CurrentScriptPath\DiaggFunctions.ps1"
#."$($Tsenv:DeployRoot)\scripts\DiaggFunctions.ps1"
Init-Function

If ($Script:OSD_Env.TSenv -eq $False){Log-ScriptEvent "[ERROR] No Tsenv: drive detected, Abroting!!!" ;  Exit}



##== Functions
function Get-RandomDate
    {
	    [DateTime]$Min = "1/1/2008"
	    [DateTime]$Max = [DateTime]::Now

	    $RandomGen = new-object random
	    $RandomTicks = [Convert]::ToInt64( ($Max.ticks * 1.0 - $Min.Ticks * 1.0 ) * $RandomGen.NextDouble() + $Min.Ticks * 1.0 )
	    $Date = new-object DateTime($RandomTicks)
	    return $Date.ToString("yyyyMMdd")
    }

##== variables
$DeployRoot = $Tsenv:DeployRoot
$LanguageToDownload = $tsenv:uilanguage
$LanguageFile = "$DeployRoot\scripts\ListOfLanguages.xml"
$productEditionId = $Tsenv:LATESTONLINEVERSION
$ProductFriendlyName = $Tsenv:LATESTONLINEVERSIONNAME
$PreferedArchitecture = 'x64'
$ISOName = $($ProductFriendlyName.Replace(" ","_") + $PreferedArchitecture + "_" + $LanguageToDownload + ".iso")
$DestinationFile = ($Tsenv:LogPath).Replace("\SMSOSD\OSDLOGS","")

$SessionId = [guid]::NewGuid()
$RequestData = @{}
$RequestData["GetLangs"] = @("a8f8f489-4c7f-463a-9ca6-5cff94d8d041", "getskuinformationbyproductedition" )
$RequestData["GetLinks"] = @("cfa9e580-a81e-4a4b-a846-7b21bf4e2e5b", "GetProductDownloadLinksBySku" )
$FirefoxVersion = Get-Random -Minimum 30 -Maximum 60
$FirefoxDate = Get-RandomDate
$UserAgent = "Mozilla/5.0 (X11; Linux i586; rv:$FirefoxVersion.0) Gecko/$FirefoxDate Firefox/$FirefoxVersion.0"

##== Get Official Language Name
[xml]$XML = Get-Content -Path $LanguageFile
$PreferedLanguage = ($XML.LOCALEDATA.LOCALE|where {$_.SSPECIFICCULTURE -eq $LanguageToDownload -and $_.IFLAGS -eq 0}).SENGDISPLAYNAME

## Get available languages
$url = "https://www.microsoft.com/en-us/api/controls/contentinclude/html"
$url += "?pageId=" + $RequestData["GetLangs"][0]
$url += "&host=www.microsoft.com"
$url += "&segments=software-download,windows10ISO" 
$url += "&query=&action=" + $RequestData["GetLangs"][1]
$url += "&sessionId=" + $SessionId
$url += "&productEditionId=" + $productEditionId
$url += "&sdVersion=2"

Log-ScriptEvent "URL: $url"

Try
    {$Reply = Invoke-WebRequest -UseBasicParsing -UserAgent $UserAgent -SessionVariable "Session" $url}
Catch
    {
        Log-ScriptEvent "[ERROR] Internet Connection Unavailable, unable to Download, Aborting" -Severity 3
        Exit
    }
[xml]$Page1 = $Reply.Content -Replace"ï»¿","" 
$Html = $page1.GetElementsByTagName('select')

Log-ScriptEvent 'Available Languages:'
foreach ($Item in $html.option) 
    {
    	$json = $Item.value | ConvertFrom-Json
		if ($json)
            {
                $LanguageList += @(New-Object PsObject -Property @{Language = $json.language; Id = $json.id })
                Log-ScriptEvent " - $($json.language) (id=$($json.id))"
            }
    }

if ($LanguageList.Length -eq 0)
    {
        Log-ScriptEvent "[ERROR] Could not parse languages, unable to Download, Aborting" -Severity 3
        Exit
    } 



## Get Download Link
$url = "https://www.microsoft.com/en-us/api/controls/contentinclude/html"
$url += "?pageId=" + $RequestData["GetLinks"][0]
$url += "&host=www.microsoft.com"
$url += "&segments=software-download,windows10ISO"
$url += "&query=&action=" + $RequestData["GetLinks"][1]
$url += "&sessionId=" + $SessionId
$url += "&skuId=" + ($LanguageList|where language -eq $PreferedLanguage).id
$url += "&language=" + $PreferedLanguage
$url += "&sdVersion=2"

Log-ScriptEvent "URL: $url"

Try
    {$Reply = Invoke-WebRequest -UseBasicParsing -UserAgent $UserAgent -WebSession $Session $url}
Catch
    {
        Log-ScriptEvent "[ERROR] Internet Connection Unavailable, unable to get Download link, Aborting" -Severity 3
        Exit
    }

[xml]$Page2 = $Reply.Content -Replace"ï»¿","" 
$Html = $page2.GetElementsByTagName('input')
$html = $($html.value -replace "IsoX86", """x86"""  -replace "IsoX64", """x64""" -replace "&amp;","&")|ConvertFrom-Json

foreach ($Item in $html) 
    {
        If ($Item.DownloadType -eq $PreferedArchitecture)
            {
                $IsoUrl = $Item.Uri
                Break
            }
    }

## Download the file
Log-ScriptEvent "Downloading ISO for $ProductFriendlyName - $PreferedLanguage"

$IsoPath = Download-File -Url $IsoUrl -Path $DestinationFile

Remove-Variable -Name LanguageList

## Mount ISO
$DriveLetter = (Mount-DiskImage -ImagePath $IsoPath | Get-Volume).DriveLetter
$DriveLetter +=  ":"
Log-ScriptEvent "Image Mounted to drive $DriveLetter"

If (test-path "$DriveLetter\sources\install.wim")
    {
        $Tsenv:MountedIsoPath = $DriveLetter
        Log-ScriptEvent "Property MountedIsoPath = $DriveLetter"
        $Tsenv:MountedIsoWIMPath = "$DriveLetter\sources\install.wim"
        Log-ScriptEvent "Property MountedIsoWIMPath = $DriveLetter\sources\install.wim"
    }
Else
    {
        Log-ScriptEvent "[ERROR] Unable to find WIM image, Aborting" -Severity 3
        Exit    
    } 

##== Get image info
$CmdLibEx = "wimlib-imagex.exe"
$CmdLib = "$DeployRoot\Tools\x64\WimLib\$CmdLibEx"
if (-not (Test-Path $CmdLib))
    {
        Log-ScriptEvent "[ERROR] unable to locate $CmdLibEx, Aborting!!!" -Severity 3
        Exit
    }


##== Extract Xml info file from WIM content
$Iret = Invoke-Executable -Path $CmdLib -Arguments "info $Tsenv:MountedIsoWIMPath --extract-xml $DestinationFile\[1].xml"


##== Find Professionnal edition's index
If ($Iret -eq 0)
    {
        [xml]$XML = Get-Content -Path (("$DestinationFile\[1].xml" -replace '\]','``]') -replace '\[','``[')
        $Tsenv:MountedIsoWIMIndex = ($xml.WIM.IMAGE| where FLAGS -eq "Professional").Index
    }
Else
    {
        Log-ScriptEvent "Unable to retrive index informations, 7zip was not succesfull, Setting WIM image index to 6" -severity 2
        $Tsenv:MountedIsoWIMIndex = 6
    }

Log-ScriptEvent "Property MountedIsoWIMIndex = $Tsenv:MountedIsoWIMIndex"

Log-ScriptEvent "************************************************************"
Log-ScriptEvent "$Script:CurrentScriptName all action finished !!!"
Log-ScriptEvent "************************************************************"