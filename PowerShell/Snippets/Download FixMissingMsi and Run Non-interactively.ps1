### Easiest method of getting the files local to a server may be to copy them from 1 host to another via admin share
# e.g. Copy-Item C:\Path\To\SQL2016 -Destination "\\ServerName\C$\Windows\Temp\SQL2016" -recurse -force
#
#

[string]$FixMissingMsiPath = "https://github.com/suyouquan/SQLSetupTools/releases/download/V2.2.1/FixMissingMSI_V2.2.1_For_NET45.zip"
[string[]]$sourcePaths     = "C:\Windows\Temp\SQL2016","C:\Windows\Temp\SQL2019","C:\Windows\Temp\SQL2017"

$ProgressPreference = "SilentlyContinue"

# download & expand .zip in $ENV:temp 
$zipPath = "$env:TEMP\FixMissingMsi.zip"
# should fix tls error for server 2016 or older hosts if PS defaults to older TLS methods
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri $FixMissingMsiPath -UseBasicParsing -OutFile $zipPath

cd $env:temp

Expand-Archive .\FixMissingMsi.zip -Force

cd FixMissingMsi

# loop each source and replace files from source using the application back-end
foreach($source in $sourcePaths){
    if(-not (Test-Path($source))){
        write-warning "$source doesnt exist, skipping"
        continue 
    }
    # Load the FixMissingMSI assembly
    $asm = [System.Reflection.Assembly]::LoadFrom("$PWD\FixMissingMSI.exe")

    # There has to be an instance of the UI form to pull a handle and call the backend methods or they will throw obj reference errors. 
    $formType = $asm.GetTypes() | Where-Object { $_.Name -eq 'Form1' }
    $form = [Activator]::CreateInstance($formType)
    # kind of hacky, well, lets be real, this whole thing is hacky but it does get the job done
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [void]$form.Handle

    # Grab the myData type & the CacheFileStatus enum
    $myDataType = $asm.GetType('FixMissingMSI.myData')
    $statusEnum = $asm.GetType('FixMissingMSI.CacheFileStatus')

    # Point it at your extracted SQL media
    $fld = $myDataType.GetField('setupSource',[Reflection.BindingFlags]'Static,Public')
    $fld.SetValue($null, $source)

    # Scan the media for all MSI/MSP packages
    $scanMedia = $myDataType.GetMethod('ScanSetupMedia',[Reflection.BindingFlags]'Static,Public')
    $scanMedia.Invoke($null, @())  # populates myData.sourcePkgs

    # Scan installed products & patches
    $scanProducts = $myDataType.GetMethod('ScanProducts',[Reflection.BindingFlags]'Static,Public')
    $scanProducts.Invoke($null, @())  # builds myData.rows with status OK/Mismatched/Missing

    # Add any extra packages from LastUsedSource folders (this mirrors what AfterDone does under the covers)
    $addFromLast = $myDataType.GetMethod('AddMsiMspPackageFromLastUsedSource',[Reflection.BindingFlags]'Static,NonPublic')
    $addFromLast.Invoke($null, @())

    # Generate the fix‐commands for each missing/mismatched row
    $updateFix = $myDataType.GetMethod('UpdateFixCommand',[Reflection.BindingFlags]'Static,Public')
    $updateFix.Invoke($null, @())  # populates each myRow.FixCommand 

    # Pull out the resulting rows
    $rowsField = $myDataType.GetField('rows',[Reflection.BindingFlags]'Static,Public')
    $rows     = $rowsField.GetValue($null)  # a SortableBindingList<myRow>

    # Filter to just missing or mismatched
    $badRows = $rows | Where-Object {$_.Status -in "Missing","Mismatched"}
    $badRowsWithFixCommand = $badRows | Where-Object {$_.FixCommand}

    # Output findings + run fix commands:
    Write-Output "$($badRows.Count) files are missing or mismatched. $($badRowsWithFixCommand.Count) missing or mismatched files will be copied from the source folder to resolve issues"

    Foreach($row in $badRowsWithFixCommand){
        & cmd /c $row.FixCommand 
    }
}
