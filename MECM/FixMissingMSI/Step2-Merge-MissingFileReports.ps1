####### Intended to be ran interactively from a relay or such to merge all the .csv reports together AFTER they have finished execution
####### Merges Report .CSV Files 
####### Don't forget to define $fileShareServer
[string]$fileShareServer = "{FILESHARESERVER}" # all computer objects must have read/write access to these paths. 

Set-Location "\\$fileShareServer\FixMissingMSI\Reports"

$csvFiles = Get-ChildItem -filter "*.csv" | Select-Object -ExpandProperty FullName

$mergedCSV = foreach($csvFile in $csvFiles){
    Import-Csv $csvFile
}

$uniqueCSV = ($mergedCSV | Sort-Object ProductCode,PackageCode,PatchCode,PackageName -Unique)

$msiProductCodes = ($uniqueCSV | Where-Object {$_.PackageName -like "*.msi"} | Select-Object ProductCode,PackageCode,Publisher,ProductVersion | Sort-Object * -unique)
$msiProductCodes | Export-CSV "MSIProductCodes.csv" -NoTypeInformation -Force

$mspPatchCodes = ($uniqueCSV | Where-Object {$_.PackageName -like "*.msp"} | Select-Object ProductCode,PatchCode,PackageName,Publisher,ProductVersion | Sort-Object * -unique)
$mspPatchCodes | Export-CSV "MSPPatchCodes.csv" -NoTypeInformation -Force
