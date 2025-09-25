####### Populate MSI and MSP Cache 
####### Intended to be deployed to all servers via MECM or other management tool
####### Using the merged .csv reports on the share, only files reported missing are copied up to the cache. 

function Compress-GUID {
    param([string]$Guid)
    $csharp = @"
using System;
public class CleanUpRegistry {
    public static string ReverseString(string s) {
        char[] a = s.ToCharArray(); Array.Reverse(a); return new string(a);
    }
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
    return [CleanUpRegistry]::CompressGUID($Guid)
}

function Get-InstalledPackageCode {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ProductCode  # e.g. {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    )

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $packageCode = $installer.ProductInfo($ProductCode, "PackageCode")
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($installer) | Out-Null
    return $packageCode
}

function Get-CachedMsiInformation {
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$ProductCode,
        [string]$DisplayName
    )

    # determine the compressed key name
    if ($ProductCode) {
        $compressed = Compress-GUID $ProductCode
    } else {
        $basePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
        $found = Get-ChildItem $basePath -ErrorAction SilentlyContinue | ForEach-Object {
            $instProps = Join-Path $_.PSPath 'InstallProperties'
            try {
                $props = Get-ItemProperty $instProps -ErrorAction Stop
                if ($props.DisplayName -eq $DisplayName) {
                    $ProductCode = $props.UninstallString.Replace("MsiExec.exe /X","")
                    return $_.PSChildName
                }
            } catch { }
        }

        if (-not $found) {
            throw "No product found with DisplayName '$DisplayName' on $ComputerName"
        }
        $compressed = $found
    }

    # read InstallProperties
    $ipPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$compressed\InstallProperties"
    $installProps = Get-ItemProperty -Path $ipPath -ErrorAction Stop

    $classesInstallerPathSourceMSI  = "HKLM:\SOFTWARE\Classes\Installer\Products\$compressed\SourceList"
    $classesInstallerPathSourceMSIProps = Get-ItemProperty -Path $classesInstallerPathSourceMSI -ErrorAction Stop

    $classesInstallerPathSourceList  = "HKLM:\SOFTWARE\Classes\Installer\Products\$compressed\SourceList\Net"
    $classesInstallerPathSourceListProps = Get-ItemProperty -Path $classesInstallerPathSourceList -ErrorAction Stop

    # compose full path
    $information = [PSCustomObject]@{
        InstallSourcePath  = $installProps.InstallSource
        CachedMsiVersion   = $installProps.DisplayVersion
        CachedMsiPath      = $installProps.LocalPackage
        CachedMsiExists    = Test-Path($installProps.LocalPackage)
        LastUsedSourcePath = $classesInstallerPathSourceListProps.1
        LastUsedSourceMsi  = $classesInstallerPathSourceMSIProps.PackageName
        ProductCode        = $ProductCode
        PackageCode        = Get-InstalledPackageCode $ProductCode
        EncodedProductCode = $compressed
    }

    return $information
}

function Get-CachedMspInformation {
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$PatchCode
    )

    # determine the compressed key name
    if ($PatchCode) {
        $compressed = Compress-GUID $PatchCode
    } else {
        $basePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Patches"
        $found = Get-ChildItem $basePath -ErrorAction SilentlyContinue | ForEach-Object {
            $instProps = Join-Path $_.PSPath 'InstallProperties'
            try {
                $props = Get-ItemProperty $instProps -ErrorAction Stop
                if ($props.DisplayName -eq $DisplayName) {
                    $PatchCode = $props.UninstallString.Replace("MsiExec.exe /X","")
                    return $_.PSChildName
                }
            } catch { }
        }

        if (-not $found) {
            throw "No product found with DisplayName '$DisplayName' on $ComputerName"
        }
        $compressed = $found
    }

    # read InstallProperties
    $ipPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Patches\$compressed"
    $installProps = Get-ItemProperty -Path $ipPath -ErrorAction Stop

    $classesInstallerPathSourceMSP  = "HKLM:\SOFTWARE\Classes\Installer\Patches\$compressed\SourceList"
    $classesInstallerPathSourceMSPProps = Get-ItemProperty -Path $classesInstallerPathSourceMSP -ErrorAction Stop

    $classesInstallerPathSourceList  = "HKLM:\SOFTWARE\Classes\Installer\Patches\$compressed\SourceList\Net"
    $classesInstallerPathSourceListProps = Get-ItemProperty -Path $classesInstallerPathSourceList -ErrorAction Stop

    # compose full path
    $information = [PSCustomObject]@{
        InstallSourcePath  = $installProps.InstallSource
        CachedMspPath      = $installProps.LocalPackage
        CachedMspExists    = Test-Path($installProps.LocalPackage)
        LastUsedSourcePath = $classesInstallerPathSourceListProps.1
        LastUsedSourceMsp  = $classesInstallerPathSourceMSPProps.PackageName
        PatchCode        = $PatchCode
        EncodedPatchCode = $compressed
    }

    return $information
}


function Get-MsiProp {
    param(
        $file
    )

    $installer = New-Object -ComObject WindowsInstaller.Installer


    if($file -like "*.msi"){
        $db = $installer.OpenDatabase($file, 0)
        $v = $db.OpenView("Select-Object `Value` FROM `Property` WHERE `Property`='ProductCode'")
        $v.Execute()

        return [PSCustomObject]@{
            ProductCode = ($v.Fetch()).StringData(1)
            PackageCode = ($installer.SummaryInformation($file,0)).Property(9)
        }
    }
}

Start-Transcript -OutputDirectory "$env:temp" | Out-Null

[string]$fileShareServer = "{FILESHARESERVER}" # all computer objects must have read/write access to these paths. 

# merged list of MSIs needed 
$CSV = Import-Csv "\\$fileShareServer\FixMissingMSI\Reports\MSIProductCodes.csv"

# check unregistered cache

[array]$RegisteredMSI = Get-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\*" | Select-Object -Expand PSPath | Foreach-Object {
    Get-ItemPropertyValue -Path "$_\InstallProperties" -Name LocalPackage
} 

$UnregisteredMSI = Get-ChildItem "C:\Windows\Installer" -Filter "*.msi" | Where-Object {$_.FullName -notin $RegisteredMSI} | Select-Object -ExpandProperty FullName 

foreach($file in $UnregisteredMSI){
    $msiProps = Get-MsiProp $file

    if($msiProps.ProductCode -in $CSV.ProductCode -and $msiProps.PackageCode -in $CSV.ProductCode){
        
        $row = $CSV | Where-Object {$_.ProductCode -eq $msiProps.ProductCode -and $_.ProductCode -eq $msiProps.PackageCode}
        
        if(-Not (Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.trim('\'))"))){
            New-Item -ItemType Directory "\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)" -Force | Out-Null
            Copy-Item $file -Destination "\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.trim('\'))" | Out-Null
            "Unregistered populated product $($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.trim('\'))"
        }
    }
}

foreach($row in $CSV){
    try{
        $information = Get-CachedMsiInformation -ProductCode $row.ProductCode
    } catch {
        continue
    }

    if($information.CachedMsiExists -eq $true -and $information.PackageCode -eq $row.PackageCode){
        if(-Not (Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)"))){
            New-Item -ItemType Directory "\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)" -Force | Out-Null
        }

        if(-Not (Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.trim('\'))"))){
           Copy-Item $information.CachedMsiPath -Destination "\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.trim('\'))" | Out-Null
           "Populated product $($row.ProductCode)\$($row.PackageCode)\$($row.PackageName.trim('\'))"
        }
    }
}

$CSV   = Import-Csv  "\\$fileShareServer\FixMissingMSI\Reports\MSPPatchCodes.csv"

foreach($row in $CSV){
    try{
        $information = Get-CachedMspInformation -PatchCode $row.PatchCode 
    } catch {
        continue
    }

    if($information.CachedMspExists -eq $true -and $information.PatchCode -eq $row.PatchCode){
        if(-Not (Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Patches\$($row.ProductCode)\$($row.PatchCode)"))){
            New-Item -ItemType Directory "\\$fileShareServer\FixMissingMSI\Cache\Patches\$($row.ProductCode)\$($row.PatchCode)" -Force | Out-Null
        }
        
        if(-Not (Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Patches\$($row.ProductCode)\$($row.PatchCode)\$($row.PackageName.trim('\'))"))){
           Copy-Item $information.CachedMspPath -Destination "\\$fileShareServer\FixMissingMSI\Cache\Patches\$($row.ProductCode)\$($row.PatchCode)\$($row.PackageName.trim('\'))" | Out-Null
           "Populated patch $($row.ProductCode)\$($row.PatchCode)\$($row.PackageName.trim('\'))"
        }
    }
}

Stop-Transcript 
