<#
.SYNOPSIS
    Synchronizes the latest Windows Server CU version strings from updates deployed via ADR in MECM to an Azure Blob storage container.
    For use with the KQL query "VMs and Arc-joined Servers Behind on Patches" in ./Azure Monitor/ of this repo
.NOTES
    Author: Joey Eckelbarger 
    Date: 07-02-2025
#>
Connect-MECM {
    # get Server Update ADRs
    $serverOSUpdateADRs = Get-CMSoftwareUpdateAutoDeploymentRule -Fast | Where {$_.Name -like "Server*Updates"}

    # get the last OS patches deployed from each including the SourceURL we can use to download from MSFT directly
    $lastUpdatesDeployed = foreach($ADR in $serverOSUpdateADRs.Name) {
        $mostRecentUpdateDeployedInADR = Get-CMSoftwareUpdate -UpdateGroupName $ADR -Fast | Where {$_.LocalizedDisplayName -notlike "*.NET Framework*" -and $_.LocalizedDisplayName -notlike "*Internet*" -and $_.LocalizedDisplayName -notlike "*Servicing Stack*"} | Sort DatePosted | Select -Last 1
        $updateDisplayName             = $mostRecentUpdateDeployedInADR | Select -ExpandProperty LocalizedDisplayName
        $updateFileUrl                 = Get-CMSoftwareUpdateContentInfo -InputObject $mostRecentUpdateDeployedInADR | Select -ExpandProperty SourceURL 

        [PSCustomObject]@{
            ADRName           = $ADR
            UpdateDisplayName = $updateDisplayName
            URL               = $updateFileUrl
        }
    }
}

# fixes performance issues w/ downloading using invoke-webrequest
$ProgressPreference = 'SilentlyContinue'

# new temp dir + set location there
$newTempDir = New-TemporaryDirectory -suffix "LatestServerPatches"
Set-Location $newTempDir

foreach($update in $lastUpdatesDeployed){
    $fileName = Split-Path $update.URL -Leaf
    $extension = $filename.split(".") | Select -last 1
    $folderName = $fileName.Replace(".$extension","")

    # download .cab update file
    Invoke-WebRequest -Uri $update.URL -OutFile $fileName 

    # make dir to extract files to 
    new-item -Type Directory -Name $folderName -Force

    # extract update.mum from the .cab 
    expand.exe $fileName -F:update.mum $folderName | out-null

    $xmlPath = "$newTempDir\$folderName\update.mum"

    # Load update.mum XML
    [xml]$xml = Get-Content $xmlPath

    # Register the mum2 namespace
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $namespaceManager.AddNamespace("mum2", "urn:schemas-microsoft-com:asm.v3")

    # select the mum2:customInformation node
    $versionNode = $xml.SelectSingleNode("//mum2:customInformation", $namespaceManager)

    # add the version string to the PSCustomObject
    $update | Add-Member -NotePropertyName "OSVersion" -NotePropertyValue $versionNode.Version -Force
}

# Clean up
Get-ChildItem -Filter "*.cab" | Remove-item -Force

# export version to file
$lastUpdatesDeployed | Select -ExpandProperty OSVersion | Out-File .\latestOSversions.txt -Encoding utf8 

# upload YAML
Connect-AzAccount -Identity -NoWelcome
$context = (Get-AzStorageAccount -Name "AUTOMATION" -ResourceGroupName "AUTOMATION").Context
Set-AzStorageBlobContent -File "latestOSversions.txt" -Container "os-patch-compliance" -Context $context -Force

Set-Location C:
Remove-Item $newTempDir -Force -Recurse
