<#
.SYNOPSIS
    Copies FixMissingMSI locally, runs it non-interactively via .NET reflection,
    attempts to source missing MSI/MSP files from local and shared caches,
    and exports a CSV report of unresolved items.

    Supports -WhatIf and -Confirm switches (high-level guard; fix execution has its own guard).

.DESCRIPTION
    Intended for broad deployment (e.g., MECM). This script stages FixMissingMSI locally,
    then loads the FixMissingMSI.exe assembly and invokes its internal methods via reflection,
    bypassing the GUI to run non-interactively.

    For each configured source path, the script:
      1) Loads FixMissingMSI.exe and instantiates the hidden Form to satisfy internal dependencies.
      2) Points FixMissingMSI to a given setup source directory (MSI/MSP media).
      3) Scans setup media, installed products/patches, and LastUsedSource locations.
      4) Generates FixCommand entries for any missing/mismatched rows.
      5) If a FixCommand is still empty, attempts to build a COPY command from the shared cache
         under \\<Server>\<Share>\FixMissingMSI\Cache\{Products,Patches}.
      6) Exports unresolved rows without a FixCommand to the central Reports folder.
      7) Executes generated FixCommands (guarded by ShouldProcess) to repopulate C:\Windows\Installer.

    > Note: FixMissingMSI is a GUI application without a native CLI. This script leverages .NET
    > reflection to call internal types and methods directly. If FixMissingMSI internals change
    > in future versions, binding calls may need to be updated.

.PARAMETER FileSharePath
    UNC to the share root that contains the app tree. Example: \\FS01\Software
    The script expects the app at: \\<Server>\<Share>\FixMissingMSI

.PARAMETER SourcePaths
    One or more setup media folders (local or UNC) to scan for MSI/MSP packages.
    Defaults to the shared Cache root: \\<Server>\<Share>\FixMissingMSI\Cache

.PARAMETER LocalWorkPath
    Local working directory where FixMissingMSI is staged and executed.
    Defaults to $env:TEMP\FixMissingMSI.

.EXAMPLE
PS> .\Step1-Invoke-FixMissingMSI.ps1 -FileSharePath \\FS01\Software

    Stages FixMissingMSI locally, scans \\FS01\Software\FixMissingMSI\Cache,
    attempts repair, exports unresolved CSV to \\FS01\Software\FixMissingMSI\Reports,
    and executes FixCommands to repopulate the local installer cache.

.EXAMPLE
PS> .\Step1-Invoke-FixMissingMSI.ps1 -FileSharePath \\FS01\Software -SourcePaths 'D:\Media','\\FS01\Builds\Office'

    Scans the provided source paths in order (D:\Media, then \\FS01\Builds\Office) instead of the default shared cache.

.NOTES
    Author: Joey Eckelbarger

    Credits:
        FixMissingMSI is authored and maintained by suyouquan
        Source: https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1

    Security:
        This script writes to C:\Windows\Installer via generated FixCommands.
        Execution of FixCommands is wrapped in ShouldProcess so -WhatIf/-Confirm can be used if desired.

    Requires:
        - PowerShell 5.1+
        - NTFS permissions to write to $LocalWorkPath
        - Read access to \\<Server>\<Share>\FixMissingMSI and subfolders
        - Write access to \\<Server>\<Share>\FixMissingMSI\Reports
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$FileSharePath,

    [Parameter()]
    [string[]]$SourcePaths = "",

    [Parameter()]
    [string]$LocalWorkPath = (Join-Path $env:TEMP 'FixMissingMSI')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Compose shared paths (app folder and caches) once. Keep all shared I/O under the app folder.
$ShareRoot      = $FileSharePath.TrimEnd('\')
$AppFolder      = Join-Path $ShareRoot 'FixMissingMSI'
$CacheRoot      = Join-Path $AppFolder 'Cache'
$ProductsCache  = Join-Path $CacheRoot 'Products'
$PatchesCache   = Join-Path $CacheRoot 'Patches'
$ReportsPath    = Join-Path $AppFolder 'Reports'

# Prepare local working directory and transcript path first (Start-Transcript requires an existing folder).
if (-not (Test-Path -LiteralPath $LocalWorkPath)) {
    New-Item -ItemType Directory -Path $LocalWorkPath -Force | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$TranscriptPath = Join-Path $LocalWorkPath "Transcript-$($env:COMPUTERNAME)-$timestamp.txt"
Start-Transcript -Path $TranscriptPath | Out-Null

try {
    # High-level guard: one decision covers staging + reflection + execution.
    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Stage FixMissingMSI, scan sources, generate and execute fix commands')) {
        return
    }

    # Stage FixMissingMSI locally. Copy only top-level binaries/config files.
    # Why: We need FixMissingMSI.exe and its dependencies, but not Cache\ or Reports\ folders.
    Get-ChildItem -Path $AppFolder -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $LocalWorkPath $_.Name) -Force
    }

    Push-Location $LocalWorkPath
    try {
        $serverName = $env:COMPUTERNAME

        foreach ($source in $SourcePaths) {
            if ($null -eq $source)) {
                Write-Warning "Scanning without a valid source may yield fewer matches."
            }

            # 1) Load the FixMissingMSI assembly
            $exePath = Join-Path $LocalWorkPath 'FixMissingMSI.exe'
            if (-not (Test-Path -LiteralPath $exePath)) {
                throw "FixMissingMSI.exe not found at expected path: $exePath"
            }
            $asm = [System.Reflection.Assembly]::LoadFrom($exePath)

            # Create an instance of the main form. The backend expects a form handle; no UI is shown.
            Add-Type -AssemblyName System.Windows.Forms
            $formType = $asm.GetTypes() | Where-Object { $_.Name -eq 'Form1' }
            if (-not $formType) { throw "Form1 type not found in FixMissingMSI assembly." }
            $form = [Activator]::CreateInstance($formType)
            [System.Windows.Forms.Application]::EnableVisualStyles()
            [void]$form.Handle  # Initialize handle to avoid null reference exceptions in backend calls.

            # 2) Access internal types and required fields
            $myDataType = $asm.GetType('FixMissingMSI.myData')
            if (-not $myDataType) { throw "Type FixMissingMSI.myData not found." }

            # Set filters off and clear filter string
            $null = $myDataType.GetField('isFilterOn', [Reflection.BindingFlags]'Static,Public')
            $fldFilterStr = $myDataType.GetField('filterString', [Reflection.BindingFlags]'Static,Public')
            if ($fldFilterStr) { $fldFilterStr.SetValue($null, '') }

            # 3) Point to setup media source directory
            $fldSetupSource = $myDataType.GetField('setupSource', [Reflection.BindingFlags]'Static,Public')
            if ($fldSetupSource) { $fldSetupSource.SetValue($null, $source) }

            # 4) Scan supplied media for MSI/MSP packages
            $scanMedia = $myDataType.GetMethod('ScanSetupMedia', [Reflection.BindingFlags]'Static,Public')
            if ($scanMedia) { $null = $scanMedia.Invoke($null, @()) }

            # 5) Scan installed products and patches
            $scanProducts = $myDataType.GetMethod('ScanProducts', [Reflection.BindingFlags]'Static,Public')
            if ($scanProducts) { $null = $scanProducts.Invoke($null, @()) }

            # 6) Include extra packages from LastUsedSource (mirrors AfterDone behavior)
            $addFromLast = $myDataType.GetMethod('AddMsiMspPackageFromLastUsedSource', [Reflection.BindingFlags]'Static,NonPublic')
            if ($addFromLast) { $null = $addFromLast.Invoke($null, @()) }

            # 7) Generate FixCommand strings for missing/mismatched rows
            $updateFix = $myDataType.GetMethod('UpdateFixCommand', [Reflection.BindingFlags]'Static,Public')
            if ($updateFix) { $null = $updateFix.Invoke($null, @()) }

            # 8) Retrieve rows collection
            $rowsField = $myDataType.GetField('rows', [Reflection.BindingFlags]'Static,Public')
            $rows = if ($rowsField) { $rowsField.GetValue($null) } else { $null }
            if (-not $rows) { Write-Verbose "No rows returned from myData.rows."; continue }

            # 9) Filter to missing or mismatched
            $badRows = $rows | Where-Object { $_.Status -in 'Missing','Mismatched' }

            # 9.5) If FixCommand is empty, try building a COPY command from the shared cache
            foreach ($row in ($badRows | Where-Object { -not $_.FixCommand })) {
                $productCandidate = Join-Path $ProductsCache (Join-Path $($row.ProductCode) (Join-Path $($row.PackageCode) $($row.PackageName)))
                $patchCandidate   = Join-Path $PatchesCache (Join-Path $($row.ProductCode) (Join-Path $($row.PatchCode)   $($row.PackageName)))

                if ($row.PackageName -and (Test-Path -LiteralPath $productCandidate)) {
                    $row.FixCommand = "COPY `"$productCandidate`" `"C:\Windows\Installer\$($row.CachedMsiMsp)`""
                    continue
                }
                if ($row.PackageName -and (Test-Path -LiteralPath $patchCandidate)) {
                    $row.FixCommand = "COPY `"$patchCandidate`" `"C:\Windows\Installer\$($row.CachedMsiMsp)`""
                    continue
                }
            }

            $badRowsWithFix    = $badRows | Where-Object { $_.FixCommand }
            $badRowsWithoutFix = $badRows | Where-Object { -not $_.FixCommand } |
                Select-Object Status, PackageName, ProductName, Publisher, LastUsedSource, InstallSource, InstallDate, ProductCode, PackageCode, PatchCode, CachedMsiMsp, CachedMsiMspVersion,
                              @{N='Hostname';E={$serverName}}, @{N='SourcePath';E={$source}}

            # Export unresolved rows to central report (per host). Overwrites by host design; adjust if you prefer timestamped files.
            $reportFile = Join-Path $ReportsPath "$serverName.csv"
            $badRowsWithoutFix | Export-Csv -Path $reportFile -NoTypeInformation -Force

            $missingCount    = ($badRows | Where-Object { $_.Status -eq 'Missing'   }).Count
            $mismatchedCount = ($badRows | Where-Object { $_.Status -eq 'Mismatched'}).Count
            Write-Output "Source: $source"
            Write-Output "Missing: $missingCount  Mismatched: $mismatchedCount  To be fixed: $($badRowsWithFix.Count)"

            # 10) Execute fix commands (guarded). Copies to C:\Windows\Installer as needed.
            foreach ($row in $badRowsWithFix) {
                if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Execute fix command for $($row.PackageName)")) {
                    & cmd /c $row.FixCommand
                }
            }
        }
    }
    finally {
        Pop-Location
    }
}
finally {
    Stop-Transcript | Out-Null
}
