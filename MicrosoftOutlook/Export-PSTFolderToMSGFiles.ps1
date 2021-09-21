#requires -version 4
<#
.SYNOPSIS
  Export PST folder to MSG files

.DESCRIPTION
  Export a root folder items from a PST file to a given folder as MSG Files
  Microsoft Outlook is required and must be running during the script execution
  PST File has be loaded into Outlook in order for the script to be able to access its content

.PARAMETER TestRun
  Switch that allows you to see what the script will create

.INPUTS
  None

.OUTPUTS
  List of files that are created

.NOTES
  Version:        1.0
  Author:         FingersOnFire
  Creation Date:  2019-08-15
  Purpose/Change: Initial script development

.EXAMPLE
  TestRun
  Export-PSTFolderToMSGFiles.ps1 -TestRun

  Normal Run
  Export-PSTFolderToMSGFiles.ps1
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (

    [Parameter(Mandatory=$false)]
    [Switch]$TestRun
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Stop
$ErrorActionPreference = 'Stop'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# The path to the PST File.
# This file should be loaded into Outlook and Outlook should be running
$PSTFile = "C:\Test\File.pst"

# The Top level folder name. Eg: Inbox, Sent Items, ....
$RootFolderName = "Sent Items"

# Export Folder Location
$ExportFolder = "C:\Temp\Export"

# Maximum Length of the subject. The remaining part will be srunk.
$MaxSubjectLength = 60

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)]
    [String]$Name
  )

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($Name -replace $re)
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Write-Host
Write-Host "************************************************"
Write-Host "Export PST Folder Items to a Folder as MSG files"
Write-Host "************************************************"
Write-Host
Write-Host "Source PST File: $($PSTFile)"
Write-Host "Root Folder to look items from: $($RootFolderName)"
Write-Host "Export folder where MSG files will be created: $($ExportFolder)"


#Create New Outlook Object
$objOutlook = New-Object -ComObject "Outlook.Application" -ErrorAction $ErrorActionPreference;
$mapi = $objOutlook.GetNamespace("mapi");

#Looking for Root Folder
$pstRootFolders = $mapi.Stores|?{($PSTFile -eq [string]$_.FilePath)}|%{$_.GetRootFolder()}

#Looking for individual emails in the selected folder
$AllEmail = $pstRootFolders.Folders|?{$_.FolderPath -match $RootFolderName}|%{$_.items}; 

Write-Host
Write-Host "Total Items Found:" $AllEmail.count
Write-Host

if($TestRun.IsPresent){
    Write-Host "THIS IS A TESTRUN. NOTHING WILL BE SAVED"
    Write-Host
}

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

foreach($Item in $AllEmail){
    $SanitizedSubject = ""

    if([string]::IsNullOrEmpty($Item.Subject)){
        $SanitizedSubject = "NoSubject"
    }
    else{
        $SanitizedSubject = Remove-InvalidFileNameChars($Item.Subject)

        # Shrink subject name if too long
        if ($SanitizedSubject.Length -gt $MaxSubjectLength) {
            $SanitizedSubject = $SanitizedSubject.Substring(0,$MaxSubjectLength) 
        }

    }

    $Filename = (get-date -Date $Item.CreationTime).ToString("yyyyMMdd-HHmm") + " " + $SanitizedSubject + ".msg"
    $FilePath = Join-Path -Path $ExportFolder -ChildPath $Filename
	
    Write-Host "$($FilePath)"

    if(-Not ($TestRun.IsPresent)){
        $Item.saveas("$($FilePath)")
    }


}