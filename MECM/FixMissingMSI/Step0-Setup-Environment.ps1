<#
.SYNOPSIS
    Prepares the FixMissingMSI automation environment: downloads the FixMissingMSI tool,
    creates the shared folder structure, stages tool binaries, and grants access to Domain Computers.

    Supports -WhatIf and -Confirm switches.

.DESCRIPTION
    Sets up the fileshare folders needed to automate FixMissingMSI.
    
    It performs the following:
    1. Downloads the FixMissingMSI zip from GitHub release.
    2. Expands the archive to a temp working folder.
    3. Creates a standardized folder layout beneath -FileSharePath (UNC):
     \\<Server>\<Share>\{FixMissingMSI, Cache\Products, Cache\Patches, Reports}
    4. Copies the FixMissingMSI binaries into \\<Server>\<Share>\FixMissingMSI.
    5. Grants the "Domain Computers" group Read and Write NTFS permissions on the share root.
    
    > Note: FixMissingMSI is a GUI application without a native CLI. Later steps invoke its 
    > internal methods via .NET reflection to run it non-interactively. This script
    > only stages the tool and prepares directories and permissions.

.PARAMETER FileSharePath
    UNC path for the share root. Example: \\FS01\FixMissingMSI

.PARAMETER FixMissingMsiUri
    URI to the FixMissingMSI zip in the upstream repository. Defaults to the current latest: V2.2.1

.PARAMETER TempPath
    Local working directory for download and extraction. Defaults to $env:TEMP.

.EXAMPLE
PS> .\Step0-Setup-Environment.ps1 -FileSharePath \\FS01\FixMissingMSI
    
    Creates the folder layout under \\FS01\FixMissingMSI, downloads and stages FixMissingMSI,
    and grants Domain Computers read/write NTFS permissions.

.NOTES
    Author: Joey Eckelbarger

    Credits:
        FixMissingMSI is authored and maintained by suyouquan
        Source: https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1
            
    Security:
        This script grants NTFS Read and Write to "Domain Computers" on the app folder only
        (\\<Server>\<Share>\FixMissingMSI). Adjust to your organization’s standards if needed.
    
    Requires:
        - PowerShell 5.1+ (for Expand-Archive)
        - Network access to the target file server
        - NTFS modify rights on the target path
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [string]$FileSharePath,

    [uri]$FixMissingMsiUri = 'https://github.com/suyouquan/SQLSetupTools/releases/download/V2.2.1/FixMissingMSI_V2.2.1_For_NET45.zip',

    [string]$TempPath = $env:TEMP
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$ProgressPreference = 'SilentlyContinue'  # Progress UI slows iwr download speed noticeably.

# Normalize and compose paths once for clarity and to avoid typos.
$ShareRoot         = $FileSharePath.TrimEnd('\')
$AppFolder         = Join-Path $ShareRoot 'FixMissingMSI'
$CacheProductsPath = Join-Path $ShareRoot 'Cache\Products'
$CachePatchesPath  = Join-Path $ShareRoot 'Cache\Patches'
$ReportsPath       = Join-Path $ShareRoot 'Reports'

$ZipPath    = Join-Path $TempPath 'FixMissingMSI.zip'
$ExpandPath = Join-Path $TempPath 'FixMissingMSI_Expanded'

# Ensure TLS 1.2 on older hosts (e.g., Server 2016) to avoid protocol negotiation failures.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# High-level guard: one decision covers the whole provisioning action.
if (-not $PSCmdlet.ShouldProcess($ShareRoot, 'Provision FixMissingMSI environment (create folders, download, copy, set ACLs)')) {
    return
}

# Clean previous temp artifacts to avoid mixing versions.
# Separate guard for destructive actions to make -Confirm meaningful here.
if (Test-Path -LiteralPath $ZipPath) {
    if ($PSCmdlet.ShouldProcess($ZipPath, 'Remove existing zip')) {
        Remove-Item -LiteralPath $ZipPath -Force
    }
}
if (Test-Path -LiteralPath $ExpandPath) {
    if ($PSCmdlet.ShouldProcess($ExpandPath, 'Remove previous expanded folder')) {
        Remove-Item -LiteralPath $ExpandPath -Recurse -Force
    }
}

# Create folder layout idempotently.
foreach ($folder in @($ShareRoot,$AppFolder,$CacheProductsPath,$CachePatchesPath,$ReportsPath)) {
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# Download upstream tool.
Invoke-WebRequest -Uri $FixMissingMsiUri -UseBasicParsing -OutFile $ZipPath

# Unblock and expand. MOTW can block execution in some environments. I don't think iwr downloads get tagged with MOTW but just to be 100% I included it.
Unblock-File -LiteralPath $ZipPath
Expand-Archive -LiteralPath $ZipPath -DestinationPath $ExpandPath -Force

# Copy tool files into $ExpandPath
Copy-Item -Path (Join-Path $ExpandPath '*') -Destination $AppFolder -Recurse -Force

# Grant Domain Computers read/write on the share root.
# targeted servers need to copy up msi/msp and write .csv reports.
$identity  = 'Domain Computers'
$rights    = [System.Security.AccessControl.FileSystemRights]'Read, Write'
$inherit   = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
$propagate = [System.Security.AccessControl.PropagationFlags]::None
$type      = [System.Security.AccessControl.AccessControlType]::Allow

$acl  = Get-Acl -LiteralPath $AppFolder
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($identity, $rights, $inherit, $propagate, $type)
$null = $acl.AddAccessRule($rule)
Set-Acl -Path $AppFolder -AclObject $acl

Write-Output "Environment setup complete."
[PSCustomObject]@{
    "Share Root"             = $ShareRoot
    "FixMissingMSI App Path" = $AppFolder
    "Cache Paths"            = @($CacheProductsPath,$CachePatchesPath) -join "`n"
    "Reports Path"           = $ReportsPath
} | Format-List 
