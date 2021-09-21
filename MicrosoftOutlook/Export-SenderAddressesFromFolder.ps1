#requires -version 4
<#
.SYNOPSIS
  Export senders email addresses from Outlook folder items

.DESCRIPTION
  This scripts goes throught the item list in a specific Outlook folder and export senders email addresses

.PARAMETER None

.INPUTS
  None

.OUTPUTS
  Attachement in a folder

.NOTES
  Version:        1.0
  Author:         FingerOnFire
  Creation Date:  31/07/2020
  Purpose/Change: Initial script development

.EXAMPLE
  Export-SenderAddressesFromOutlookFolder.ps1
  
  Show All senders email address (1 per line)
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
Write-Host "Exporting email addresses from" $folder.Name 

foreach ($item in $folder.Items) {
	Write-Host $item.Sender.Address
}