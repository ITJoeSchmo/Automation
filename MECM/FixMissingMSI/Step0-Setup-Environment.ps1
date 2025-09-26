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
    3. Creates a standardized layout beneath -FileSharePath\FixMissingMSI:
       \\<Server>\<Share>\FixMissingMSI\
          Cache\Products\
          Cache\Patches\
          Reports\
    4. Copies the FixMissingMSI binaries into \\<Server>\<Share>\FixMissingMSI.
    5. Grants NTFS permissions:
       - App folder (FixMissingMSI): "Domain Computers" = Read & Execute
       - Cache and Reports:          "Domain Computers" = Read, Write (CI/OI)
    
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
PS> .\Step0-Setup-Environment.ps1 -FileSharePath "\\FS01\"
    
    Creates \\FS01\Software\FixMissingMSI with the required subfolders, downloads and stages FixMissingMSI,
    sets read/execute on the app folder and read/write on Cache and Reports for Domain Computers.

.NOTES
    Author: Joey Eckelbarger

    Credits:
        FixMissingMSI is authored and maintained by suyouquan
        Source: https://github.com/suyouquan/SQLSetupTools/releases/tag/V2.2.1
            
    Security:
        App folder is readable but not writable by "Domain Computers".
        Cache and Reports are writable to allow servers to upload MSI/MSP files and CSV reports.
    
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
$CacheRoot         = Join-Path $AppFolder 'Cache'
$CacheProductsPath = Join-Path $CacheRoot 'Products'
$CachePatchesPath  = Join-Path $CacheRoot 'Patches'
$ReportsPath       = Join-Path $AppFolder 'Reports'

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
Try {
    foreach ($folder in @($AppFolder,$CacheProductsPath,$CachePatchesPath,$ReportsPath)) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
    }
} Catch {
    Write-Error $_
    Throw "Failed to create folders on $FileSharePath"
}

# Download upstream tool.
Invoke-WebRequest -Uri $FixMissingMsiUri -UseBasicParsing -OutFile $ZipPath

# Unblock and expand. MOTW can block execution in some environments. I don't think iwr downloads get tagged with MOTW but just to be 100% I included it.
Unblock-File -LiteralPath $ZipPath
Expand-Archive -LiteralPath $ZipPath -DestinationPath $ExpandPath -Force

# Copy tool files into $ExpandPath
Copy-Item -Path (Join-Path $ExpandPath '*') -Destination $AppFolder -Recurse -Force

# Identity and rights
# targeted servers need to copy up msi/msp and write .csv reports.
$domainComputers = 'Domain Computers'
$readExec = [System.Security.AccessControl.FileSystemRights]'ReadAndExecute'
$readWriteDelete = [System.Security.AccessControl.FileSystemRights]::Modify
$ci = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit'
$oi = [System.Security.AccessControl.InheritanceFlags]'ObjectInherit'
$inheritBoth = $ci -bor $oi
$none = [System.Security.AccessControl.PropagationFlags]::None
$allow = [System.Security.AccessControl.AccessControlType]::Allow

# 1) App folder: ensure Domain Computers have Read & Execute only.
#    Break inheritance on the app folder so any parent write permissions do not flow down.
$aclApp = Get-Acl -LiteralPath $AppFolder
$aclApp.SetAccessRuleProtection($true, $true)  # protect; convert inherited to explicit
$aceReadExec = New-Object System.Security.AccessControl.FileSystemAccessRule($domainComputers, $readExec, $inheritBoth, $none, $allow)
[void]$aclApp.AddAccessRule($aceReadExec)
Set-Acl -Path $AppFolder -AclObject $aclApp

# 2) Cache: Domain Computers Read + Write (CI/OI)
$aclCache = Get-Acl -LiteralPath $CacheRoot
$aceRW = New-Object System.Security.AccessControl.FileSystemAccessRule($domainComputers, $readWriteDelete, $inheritBoth, $none, $allow)
[void]$aclCache.AddAccessRule($aceRW)
Set-Acl -Path $CacheRoot -AclObject $aclCache

# 3) Reports: Domain Computers Read + Write (CI/OI)
$aclReports = Get-Acl -LiteralPath $ReportsPath
[void]$aclReports.AddAccessRule($aceRW)
Set-Acl -Path $ReportsPath -AclObject $aclReports

Write-Output "Environment setup complete."
[PSCustomObject]@{
    "Share Root"             = $ShareRoot
    "FixMissingMSI App Path" = $AppFolder
    "Reports Path"           = $ReportsPath
    "Cache Paths"            = @($CacheProductsPath,$CachePatchesPath) -join "`n"
} | Format-List 
