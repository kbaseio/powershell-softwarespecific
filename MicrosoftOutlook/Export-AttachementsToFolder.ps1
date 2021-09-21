#requires -version 4
<#
.SYNOPSIS
  Export attachement from Outlook folder items

.DESCRIPTION
  This scripts goes throught the item list in a specific Outlook folder and export all attachement to folder on your harddrive

.PARAMETER None

.INPUTS
  None

.OUTPUTS
  Attachement in a folder

.NOTES
  Version:        1.0
  Author:         FingerOnFire
  Creation Date:  03/07/2020
  Purpose/Change: Initial script development

.EXAMPLE
  Export-AttachementsToFolder.ps1
  
  Export all attachement based on the golbal paramter of the script
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  #Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

$ErrorActionPreference = 'SilentlyContinue'


#Import Modules & Snap-ins
#Nothing Here

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$filepath = "c:\Temp\Attachements\"

# Date Range for the attachement export. Format is US : MM/DD/YY
$dateStart = [datetime]"01/01/20"
$dateEnd = [datetime]"06/30/20"

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#set outlook to open
$Outlook = New-Object -ComObject "Outlook.Application" -ErrorAction $ErrorActionPreference;
$MAPI = $Outlook.GetNamespace("MAPI");

#you'll get a popup in outlook at this point where you pick the folder you want to scan
Write-Host
Write-Host "Please select the folder from Outlook window"
$folder = $MAPI.pickfolder()

#loop through items from the selected folder and grab the attachments
Write-Host
Write-Host "Exporting Attachements from" $folder.Name 

foreach ($item in $folder.Items) {
    if(($item.ReceivedTime -gt $dateStart) -and ($item.ReceivedTime -lt $dateEnd)){

        foreach ($attachment in $item.attachments) {
            Write-Host $attachment.filename
            $attachment.saveasfile((Join-Path $filepath $attachment.filename))
        }
    }
}