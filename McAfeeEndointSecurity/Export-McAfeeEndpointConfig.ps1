#Requires -RunAsAdministrator

# ACCESS DENIED WHEN RUNNING ESConfigTool.exe
# Standalone Installation : Please temporarly disable Access Protection under Threat Prevention settings
# ePo : Please enable the following policy "ePO > Policy catalog > Endpoint Security Threat Prevention : Policy Category > Access Protection > Select the Applied policy > Rules > Unauthorized execution of EsConfigTool"
# Source : https://community.mcafee.com/t5/Endpoint-Security-ENS/ESconfigTool-Not-generating-any-output-screen-or-file/td-p/668892

$EsConfigToolPath = "C:\Program Files\McAfee\Endpoint Security\Endpoint Security Platform\ESConfigTool.exe"
$OutputFolderPath = "C:\McAfeeConfigExport\"
$OutputFilePrefix = "McAfeeEndointSecurity"
$OutputFileExtension = ".policy"
$Modules = @('TP','FW','WC', 'ATP','ESP')
$Date = Get-Date -Format "yyyyMMdd"
$Hostname = $env:COMPUTERNAME

# Create Output Directory if it does not exist
if(-Not (Test-Path -Path $OutputFolderPath)){
    New-Item -ItemType Directory -Path $OutputFolderPath -Force
}

# Export configuration for each module
foreach($Module in $Modules){

    $OutputFileName ="$OutputFilePrefix-$Hostname-$Date-$Module$OutputFileExtension"
    $OutputFilePath = Join-Path -Path $OutputFolderPath -ChildPath $OutputFileName
    $ArgumentList = "/export $OutputFilePath /module $Module"
        
    Write-Host "Exporting module :" $Module

    Start-Process -FilePath $EsConfigToolPath -ArgumentList $ArgumentList -NoNewWindow -Wait
}