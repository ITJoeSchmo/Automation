####### Step 0 : Environment Setup
[string]$fileShareServer   = "{FILESHARESERVER}"
[string]$FixMissingMsiPath = "https://github.com/suyouquan/SQLSetupTools/releases/download/V2.2.1/FixMissingMSI_V2.2.1_For_NET45.zip"

# this UI slows downloads signficantly 
$ProgressPreference = "SilentlyContinue"

# download & expand .zip in $ENV:temp 
$zipPath = "$env:TEMP\FixMissingMsi.zip"
# should fix tls error for server 2016 or older hosts if PS defaults to older TLS methods
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Invoke-WebRequest -Uri $FixMissingMsiPath -UseBasicParsing -OutFile $zipPath

Set-Location $env:temp

# unblock if it has MOTW
Unblock-File .\FixMissingMsi.zip
Expand-Archive .\FixMissingMsi.zip -Force


# make folder on fileshare
New-Item -ItemType Directory "$fileShareServer\FixMissingMSI" -Force | out-null

Get-ChildItem "\\$fileShareServer\FixMissingMSI" -File | Foreach-Object {
    Copy-Item $_.FullName -Destination "\\$fileShareServer\FixMissingMSI\$_" -Force | out-null
}

# grant read/write to all domain computer objects 
# get the current ACL
$acl = Get-Acl "\\$fileShareServer\FixMissingMSI"

# create the access rule:
# - Read + Write permissions
# - Applies to folder, subfolders, and files
# - No special propagation restrictions
# - Allow rule
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "Domain Computers",
    "Read, Write",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)

# add the rule
$acl.AddAccessRule($rule)

# update security share settings 
Set-Acl -Path "\\$fileShareServer\FixMissingMSI" -AclObject $acl
