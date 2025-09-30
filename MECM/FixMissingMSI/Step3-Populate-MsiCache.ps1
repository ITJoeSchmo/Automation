<#
.SYNOPSIS
    Populates the shared MSI/MSP cache from local machines using merged report CSVs.

.DESCRIPTION
    Intended to be deployed via MECM (or similar) after Step2-Merge-MissingFileReports.ps1 has produced:
        - MSIProductCodes.csv
        - MSPPatchCodes.csv

    This script:
      1. Reads the merged MSI/MSP summary CSVs from \\<Server>\<Share>\FixMissingMSI\Reports.
      2. Identifies any "unregistered" MSI files present in C:\Windows\Installer that are not listed
         as LocalPackage for any installed product, extracts ProductCode/PackageCode, and uploads them.
      3. For each MSI in the report, queries local registry/Windows Installer to locate the cached MSI
         and uploads it to the shared cache (if not already present).
      4. Repeats similar logic for MSP patch files.

    > Note: Only files referenced in the merged CSVs are considered for upload to the cache.
    > This minimizes noise and ensures we only collect files that were reported missing elsewhere.

.PARAMETER FileSharePath
    UNC path to the share root hosting the FixMissingMSI app folder.
    Example: \\FS01\Software
    The script expects:
      \\<Server>\<Share>\FixMissingMSI\Reports
      \\<Server>\<Share>\FixMissingMSI\Cache\{Products,Patches}

.PARAMETER LocalTranscriptPath
    Optional path for the transcript file. Defaults to $env:TEMP\FixMissingMSI\Transcript-Step3-<host>-<timestamp>.txt

.EXAMPLE
PS> .\Step3-Populate-MsiCache.ps1 -FileSharePath \\FS01\Software

    Uploads locally cached MSI/MSP files to \\FS01\Software\FixMissingMSI\Cache
    based on the merged reports in \\FS01\Software\FixMissingMSI\Reports.

.NOTES
    Author: Joey Eckelbarger

    Credits:
        The Compress-GUID implementation is adapted from Microsoft’s
        "Windows Program Install and Uninstall Troubleshooter" logic and reworked for this script.

    Requires:
        - PowerShell 5.1+
        - Read access to \\<Server>\<Share>\FixMissingMSI\Reports
        - Write access to \\<Server>\<Share>\FixMissingMSI\Cache\Products and \Cache\Patches
        - Local permission to query HKLM:\...\Installer and COM WindowsInstaller.Installer

#>
param(
    [Parameter(Mandatory = $true)]
    [string]$FileSharePath
)

$ErrorActionPreference = 'Stop'

# Compose shared paths once; all shared I/O happens under the FixMissingMSI app tree.
$ShareRoot      = $FileSharePath.TrimEnd('\')
$AppFolder      = Join-Path $ShareRoot 'FixMissingMSI'
$ReportsPath    = Join-Path $AppFolder 'Reports'
$CacheRoot      = Join-Path $AppFolder 'Cache'
$ProductsCache  = Join-Path $CacheRoot 'Products'
$PatchesCache   = Join-Path $CacheRoot 'Patches'

# Prepare transcript path and ensure folder exists.
$LocalWork = Join-Path $env:TEMP 'FixMissingMSI'
if (-not (Test-Path -LiteralPath $LocalWork)) {
    New-Item -ItemType Directory -Path $LocalWork -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$LocalTranscriptPath = Join-Path $LocalWork "Transcript-Step3-$($env:COMPUTERNAME)-$timestamp.txt"

Start-Transcript -Path $LocalTranscriptPath | Out-Null

<#
.SYNOPSIS
    Returns the compressed form of a ProductCode GUID (e.g. used in HKLM\...\Installer 
    keys to identify Products rather than the typical {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} GUID).
    
.NOTES
    Compress-GUID implementation adapted from Microsoft’s
    "Program Install and Uninstall Troubleshooter" tool.

    This reimplementation is provided as-is for compatibility with Windows Installer.
    Microsoft retains copyright to its original tool; this script is an independent
    rework that replicates the same transformation logic.
#>
function Compress-GUID {
    param([Parameter(Mandatory=$true)][string]$Guid)

    $csharp = @"
using System;
public class CleanUpRegistry {
    public static string ReverseString(string s) { char[] a = s.ToCharArray(); Array.Reverse(a); return new string(a); }
    public static string CompressGUID(string g) {
        g = g.Substring(1,36);
        return ReverseString(g.Substring(0,8)) +
               ReverseString(g.Substring(9,4)) +
               ReverseString(g.Substring(14,4)) +
               ReverseString(g.Substring(19,2)) +
               ReverseString(g.Substring(21,2)) +
               ReverseString(g.Substring(24,2)) +
               ReverseString(g.Substring(26,2)) +
               ReverseString(g.Substring(28,2)) +
               ReverseString(g.Substring(30,2)) +
               ReverseString(g.Substring(32,2)) +
               ReverseString(g.Substring(34,2));
    }
}
"@
    if (-not [Type]::GetType('CleanUpRegistry')) {
        Add-Type -TypeDefinition $csharp -Language CSharp
    }
    [CleanUpRegistry]::CompressGUID($Guid)
}

function Get-InstalledPackageCode {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ProductCode  # {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    )
    # Windows Installer exposes package metadata via the WindowsInstaller.Installer COM interface.
    $installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        $installer.ProductInfo($ProductCode, 'PackageCode')
    }
    finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
    }
}

function Get-CachedMsiInformation {
    param(
        [string]$ProductCode,
        [string]$DisplayName
    )
    # Determine compressed product key (as used in HKLM\...\Installer).
    if ($ProductCode) {
        $compressed = Compress-GUID $ProductCode
    } else {
        $basePath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products'
        $found = Get-ChildItem $basePath -ErrorAction SilentlyContinue | ForEach-Object {
            $instProps = Join-Path $_.PSPath 'InstallProperties'
            $props = Get-ItemProperty $instProps -ErrorAction Ignore
            if ($props.DisplayName -eq $DisplayName) {
                $_.PSChildName
            }
        }
        if (-not $found) { throw "No product found with DisplayName '$DisplayName'" }
        $compressed = $found
    }

    # Read properties from Installer hives.
    $ipPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$compressed\InstallProperties"
    $installProps = Get-ItemProperty -Path $ipPath -ErrorAction Stop

    $classesSourceMSI      = "HKLM:\SOFTWARE\Classes\Installer\Products\$compressed\SourceList"
    $classesSourceMSIProps = Get-ItemProperty -Path $classesSourceMSI -ErrorAction SilentlyContinue
    $classesSourceNet      = "HKLM:\SOFTWARE\Classes\Installer\Products\$compressed\SourceList\Net"
    $classesSourceNetProps = Get-ItemProperty -Path $classesSourceNet -ErrorAction SilentlyContinue

    [PSCustomObject]@{
        InstallSourcePath  = $installProps.InstallSource
        CachedMsiVersion   = $installProps.DisplayVersion
        CachedMsiPath      = $installProps.LocalPackage
        CachedMsiExists    = [bool](Test-Path -LiteralPath $installProps.LocalPackage)
        LastUsedSourcePath = $classesSourceNetProps.'1'
        LastUsedSourceMsi  = $classesSourceMSIProps.PackageName
        ProductCode        = $ProductCode
        PackageCode        = (Get-InstalledPackageCode -ProductCode $ProductCode)
        EncodedProductCode = $compressed
    }
}

function Get-CachedMspInformation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PatchCode
    )
    
    $compressed = Compress-GUID $PatchCode

    $patchRoot = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Patches\$compressed"
    $installProps = Get-ItemProperty -Path $patchRoot -ErrorAction Stop

    $classesSourceMSP      = "HKLM:\SOFTWARE\Classes\Installer\Patches\$compressed\SourceList"
    $classesSourceMSPProps = Get-ItemProperty -Path $classesSourceMSP -ErrorAction SilentlyContinue
    $classesSourceNet      = "HKLM:\SOFTWARE\Classes\Installer\Patches\$compressed\SourceList\Net"
    $classesSourceNetProps = Get-ItemProperty -Path $classesSourceNet -ErrorAction SilentlyContinue

    if($installProps.PSObject.Properties.Name -notcontains "InstallSource"){
        $calculatedInstallSource = (Join-Path $classesSourceNetProps.'1' $classesSourceMSPProps.PackageName)
        if($calculatedInstallSource){
            $installProps | Add-Member -NotePropertyName "InstallSource" -NotePropertyValue $calculatedInstallSource
        } else {
            $installProps | Add-Member -NotePropertyName "InstallSource" -NotePropertyValue ""
        }
    }

    [PSCustomObject]@{
        InstallSourcePath  = $installProps.InstallSource
        CachedMspPath      = $installProps.LocalPackage
        CachedMspExists    = [bool](Test-Path -LiteralPath $installProps.LocalPackage)
        LastUsedSourcePath = $classesSourceNetProps.'1'
        LastUsedSourceMsp  = $classesSourceMSPProps.PackageName
        PatchCode          = $PatchCode
        EncodedPatchCode   = $compressed
    }
}

<#
.SYNOPSIS
    Reads ProductCode and PackageCode from an MSI file.
.DESCRIPTION
    Uses Windows Installer COM to open the MSI database and SummaryInformation.
    ProductCode is read from the Property table; PackageCode is the SummaryInformation revision GUID.
#>
function Get-MsiProp {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )
    $installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        if ($Path -like '*.msi') {
            $db = $installer.OpenDatabase($Path, 0) # 0 = read-only
            $view = $db.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductCode'")
            $view.Execute() | Out-Null
            $rec = $view.Fetch()
            $productCode = if ($rec) { $rec.StringData(1) } else { $null }
            $pkgCode = ($installer.SummaryInformation($Path,0)).Property(9) # PID_REVNUMBER (PackageCode)
            [PSCustomObject]@{
                ProductCode = $productCode
                PackageCode = $pkgCode
            }
        }
    }
    finally {
        $null = $view.Close() 
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($view)
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
    }
}

# Load merged MSI report
$msiReport = Join-Path $ReportsPath 'MSIProductCodes.csv'
if (-not (Test-Path -LiteralPath $msiReport)) {
    Write-Warning "MSI report not found: $msiReport. Ensure host can access the fileshare."
} else {
    $MSIList = Import-Csv -LiteralPath $msiReport
    
    # Build list of currently registered LocalPackage MSIs (to identify "unregistered" MSI files).
    $registeredLocal = Get-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\*' |
                       Select-Object -ExpandProperty PSPath |
                       ForEach-Object { Get-ItemPropertyValue -Path "$_\InstallProperties" -Name LocalPackage -ErrorAction Ignore } |
                       Where-Object { $_ } |
                       Select-Object -Unique
    
    # Unregistered MSIs under C:\Windows\Installer that are not listed as LocalPackage.
    $unregistered = Get-ChildItem 'C:\Windows\Installer' -Filter '*.msi' -File |
                    Where-Object { $_.FullName -notin $registeredLocal } |
                    Select-Object -ExpandProperty FullName
    
    # Upload any unregistered MSIs that match the merged report identity.
    foreach ($file in $unregistered) {
        $props = Get-MsiProp -Path $file
        if (-not $props -or -not $props.ProductCode -or -not $props.PackageCode) { continue }
    
        $row = $MSIList | Where-Object { $_.ProductCode -eq $props.ProductCode -and $_.PackageCode -eq $props.PackageCode } | Select-Object -First 1
    
        # skip if this isn't in the missing report
        if (-not $row) { continue }
    
        $destDir  = Join-Path (Join-Path $ProductsCache $row.ProductCode) $row.PackageCode
        $destFile = Join-Path $destDir ($row.PackageName.Trim('\'))
    
        if (-not (Test-Path -LiteralPath $destFile)) {
            if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
            Copy-Item -LiteralPath $file -Destination $destFile -Force
            "Unregistered populated product $($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.Trim('\'))"
        }
    }
    
    # Upload registered/cached MSIs referenced by the merged report.
    foreach ($row in $MSIList) {
        try {
            $info = Get-CachedMsiInformation -ProductCode $row.ProductCode
        } catch { continue }
    
        if ($info.CachedMsiExists -and $info.PackageCode -eq $row.PackageCode) {
            $destDir  = Join-Path (Join-Path $ProductsCache $row.ProductCode) $row.PackageCode
            $destFile = Join-Path $destDir ($row.PackageName.Trim('\'))
    
            if (-not (Test-Path -LiteralPath $destDir))  { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
            if (-not (Test-Path -LiteralPath $destFile)) {
                Copy-Item -LiteralPath $info.CachedMsiPath -Destination $destFile -Force
                "Populated product $destFile"
            }
        }
    }
}
# Load merged MSP report
$mspReport = Join-Path $ReportsPath 'MSPPatchCodes.csv'
if (-not (Test-Path -LiteralPath $mspReport)) {
    Write-Warning "MSP report not found: $mspReport. Ensure host can access $FileSharePath"
} else {
    $MSPList = Import-Csv -LiteralPath $mspReport
    
    foreach ($row in $MSPList) {
        try {
            $info = Get-CachedMspInformation -PatchCode $row.PatchCode
        
    
            if ($info.CachedMspExists -and $info.PatchCode -eq $row.PatchCode) {
                $destDir  = Join-Path (Join-Path $PatchesCache $row.ProductCode) $row.PatchCode
                $destFile = Join-Path $destDir ($row.PackageName.Trim('\'))
    
                if (-not (Test-Path -LiteralPath $destDir))  { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                if (-not (Test-Path -LiteralPath $destFile)) {
                    Copy-Item -LiteralPath $info.CachedMspPath -Destination $destFile -Force
                    "Populated patch $destFile"
                }
            }
        } catch { continue }
    }
}
Stop-Transcript | Out-Null
