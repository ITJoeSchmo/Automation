<#
.SYNOPSIS
    Merges per-host FixMissingMSI report CSVs into 1 consolidated report that is leveraged to build the shared cache of missing files.

.DESCRIPTION
    Intended to be run interactively after Step1-Invoke-FixMissingMSI.ps1 has been deployed across servers.
    This script reads all CSV files from:
        \\<Server>\<Share>\FixMissingMSI\Reports
    and produces consolidated outputs in the same folder.

    It performs the following:
      1. Imports all *.csv under the Reports folder.
      2. Deduplicates rows by key fields (ProductCode, {PackageCode, PatchCode}, PackageName, Publisher, ProductVersion)
      3. Produces MSI and MSP summary lists:
         - MSIProductCodes.csv  (ProductCode, PackageCode, PackageName, Publisher, ProductVersion)
         - MSPPatchCodes.csv    (ProductCode, PatchCode,  PackageName, Publisher, ProductVersion)

    > Note: Step 1 adds Hostname or SourcePath columns for traceability. Those are preserved in the merged data
    > but are not part of the uniqueness key. Adjust the Select-Object properties if you prefer a different definition.

.PARAMETER FileSharePath
    UNC to the share root that contains the app tree. Example: \\FS01\Software
    The script expects reports at: \\<Server>\<Share>\FixMissingMSI\Reports

.EXAMPLE
PS> .\Step2-Merge-MissingFileReports.ps1 -FileSharePath \\FS01\Software

    Merges all host reports from \\FS01\Software\FixMissingMSI\Reports and writes
    MSIProductCodes.csv and MSPPatchCodes.csv in that folder.

.NOTES
    Author: Joey Eckelbarger

    Requires:
        - PowerShell 5.1+
        - Read/Write access to \\<Server>\<Share>\FixMissingMSI\Reports
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$FileSharePath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Compose paths; all activity occurs under the app folder's Reports directory.
$ShareRoot     = $FileSharePath.TrimEnd('\')
$AppFolder     = Join-Path $ShareRoot 'FixMissingMSI'
$ReportsPath   = Join-Path $AppFolder 'Reports'
$msiReportPath = Join-Path $ReportsPath "MSIProductCodes.csv"
$mspReportPath = Join-Path $ReportsPath "MSPPatchCodes.csv"

if (-not (Test-Path -LiteralPath $ReportsPath)) {
    throw "Reports folder not found: $ReportsPath"
}

# Discover CSVs first; fail gracefully if none are present.
[array]$csvFiles = Get-ChildItem -Path $ReportsPath -Filter '*.csv' | Where-object {$_.Name -notin @("MSIProductCodes.csv","MSPPatchCodes.csv")} | Select-Object -ExpandProperty FullName

if (-not $csvFiles -or $csvFiles.Count -eq 0) {
    Throw "No CSV reports found in $ReportsPath. Ensure Step 1 has completed across targets."
}

# Import all CSVs into 1 merged var
$merged = foreach ($csv in $csvFiles) {
    Import-Csv -LiteralPath $csv
}

# Build MSI summary (msi packages typically lack PatchCode; keep fields most relevant to sourcing).
$msi = $merged |
    Where-Object { $_.PackageName -like '*.msi' } |
    Select-Object ProductCode, PackageCode, PackageName, Publisher, ProductVersion |
    Sort-Object * -Unique

# Build MSP summary (patches carry PatchCode).
$msp = $merged |
    Where-Object { $_.PackageName -like '*.msp' } |
    Select-Object ProductCode, PatchCode, PackageName, Publisher, ProductVersion |
    Sort-Object * -Unique
if($msi){
    $msi | Export-Csv -LiteralPath $msiReportPath -NoTypeInformation -Force
} else {
    $msi = @()
}

if($msp){
    $msp | Export-Csv -LiteralPath $mspReportPath -NoTypeInformation -Force
} else {
    $msp = @()
}

# Simple summary 
[PSCustomObject]@{
    "CSV Files" = $($csvFiles.Count)
    "MSI"       = "$($msi.Count)  -> MSIPatchCodes.csv"
    "MSP"       = "$($msp.Count)  -> MSPProductCodes.csv"
} | Format-List
