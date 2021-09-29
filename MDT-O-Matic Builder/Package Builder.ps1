param(
        [parameter(Mandatory = $false)][string]$ScriptFile = "MDT-O-Matic-Template-v5.ps1", 
        [parameter(Mandatory = $false)][string]$FolderToAdd = "EmbeddedContent", 
        [parameter(Mandatory = $false)][string]$InstallTag = "##EMBED-HERE##", 
        [parameter(Mandatory = $false)][string]$DSSource = "E:\Deploy"     
    )


#########################

# Package Builder for MDT-O-Matic
# Embed Archive in your powershell script  

# By Diagg/OSDC

# Version 1.1 - Release Date 09.27.2021 - Added support for Powershell 7
#                                       - Final file is saved outside of the folder. 

#########################

<# decoding routine

##== Decoding
$Content = [System.Convert]::FromBase64String($ArchivedContent)
Set-Content -Path "$env:temp\Embedded.zip" -Value $Content -Encoding Byte

##== unzipping
$Hangar18 = "$Script:CurrentScriptPath\Extracted"
Expand-Archive -LiteralPath "$env:temp\Embedded.zip" -DestinationPath $Hangar18 -Force
Remove-Item "$env:temp\Embedded.zip"

#>


#Requires -Version 4
##Requires -RunAsAdministrator 

##== Debug
$ErrorActionPreference = "stop"
#$ErrorActionPreference = "Continue"
$Error.Clear()

##== Global Variables
$Script:CurrentScriptName = $MyInvocation.MyCommand.Name
$Script:CurrentScriptFullName = $MyInvocation.MyCommand.Path
$Script:CurrentScriptPath = split-path $MyInvocation.MyCommand.Path


##== Check If Script exists
If ($ScriptFile.Contains(":"))
    {$FileStatus = Test-path $ScriptFile}
Else
    {
        $FileStatus = test-path "$CurrentScriptPath\$ScriptFile"
        If ($FileStatus){$ScriptFile = "$CurrentScriptPath\$ScriptFile"}
    }

If (-not $FileStatus){Write-Error "[ERROR] Script file not found, aborting!!!";Exit}


##== Check If folder exists
If ($FolderToAdd.Contains(":"))
    {$FileStatus = Test-path $FolderToAdd}
Else
    {
        $FileStatus = test-path "$CurrentScriptPath\$FolderToAdd"
        If ($FileStatus){$FolderToAdd = "$CurrentScriptPath\$FolderToAdd"}
    }

If (-not $FileStatus){Write-Error "[ERROR] Script file not found, aborting!!!";Exit}


##== Fill in folder
##== this part is here for convenience adn should no be part of this script... but time is running out!!!!

##==Copy From Scripts folder
$Files = [PSCustomObject]@{ Source = "$DSSource\Scripts" ; Destination = $FolderToAdd ; FileList = @("DeployWiz_LanguageUIOnline.vbs","DeployWiz_LanguageUIOnline.xml","DiaggFunctions.ps1","Download-Iso.ps1")}
ForEach ($item in $Files.FileList){Copy-Item -Path "$($files.source)\$item" -Destination "$($files.Destination)\$item" -Force}

##==Copy From Operating Sytem folder
Copy-Item -Path "$DSSource\Operating Systems\Windows 10 Latest Online\Install.wim" -Destination "$FolderToAdd\Install.wim" -Force


##== Check if Install Tag can be found in the script file
$content = Get-Content -Path $ScriptFile

If (-not ($content| Where-Object { $_.Contains($InstallTag)}))
    {Write-Error "[ERROR] InstallTag $InstallTag no found in file $ScriptFile, aborting!!!";Exit}
Else
    {Write-Host "Tag $InstallTag successfully detected on file $ScriptFile"}

##== Create archive with associated folder
$TmpFile = "$env:temp\TempArchive.zip"
Write-Host "Temporary archive file set to ""$($TmpFile)""."
Compress-Archive -Path "$FolderToAdd\*" -DestinationPath $TmpFile -force

##== Convert Archive to base 64
$ArchiveContent = Get-Content -Path $TmpFile -AsByteStream
$Base64 = [System.Convert]::ToBase64String($ArchiveContent)
Remove-Item $TmpFile -Force

##== commit content to destination file
$NewFileName = split-path $scriptFile -leaf
$StringToReplace = $NewFileName.replace("MDT-O-Matic","")
$NewFilePath = "$(split-path $Script:CurrentScriptPath)\$($NewFileName.Replace($StringToReplace,".ps1"))"

$newContent = $content -replace $InstallTag, $base64
$newContent | Set-Content -Path $NewFilePath -Force -Encoding UTF8
Write-Host "File saved to $NewFilePath"