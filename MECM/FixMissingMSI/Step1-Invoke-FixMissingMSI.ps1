####### Copy FixMissingMSI Locally & Run, Copy MSIs from local and shared cache if available, and Report Files That Can't Be Sourced
####### Intended to be deployed to all servers via MECM or other management tool. 
####### Don't forget to define $fileShareServer
[string]$fileShareServer = "{FILESHARESERVER}" # all computer objects must have read/write access to these paths. 
[string[]]$sourcePaths    = "\\$fileShareServer\FixMissingMSI\Cache" # can be a array of strings

Start-Transcript -OutputDirectory "$env:temp\FixMissingMSI"  | Out-Null

$servername = $env:computername

New-Item -ItemType Directory "$env:temp\FixMissingMSI" -Force | out-null

Get-ChildItem "\\$fileShareServer\FixMissingMSI" -File | Foreach-Object {
    Copy-Item $_.FullName -Destination "$env:temp\FixMissingMSI\$_" -Force | out-null
}

Set-Location "$env:temp\FixMissingMSI"

foreach($source in $sourcePaths){
    if(-not (Test-Path($source))){
        write-warning "$source doesnt exist, not scanning against anything specific may yield less results..."
    }
    # 1) Load the FixMissingMSI assembly
    $asm = [System.Reflection.Assembly]::LoadFrom("$PWD\FixMissingMSI.exe")

    # There has to be an instance of the UI form to pull a handle and call the backend methods or they will throw obj reference errors. 
    $formType = $asm.GetTypes() | Where-Object { $_.Name -eq 'Form1' }
    $form = [Activator]::CreateInstance($formType)
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [void]$form.Handle

    # 2) Grab the myData type & the CacheFileStatus enum
    $myDataType = $asm.GetType('FixMissingMSI.myData')
    $statusEnum = $asm.GetType('FixMissingMSI.CacheFileStatus')

    $fldFilterOn = $myDataType.GetField(
      'isFilterOn',
      [Reflection.BindingFlags] 'Static,Public'
    )

    # 3) Point it at your extracted media
    $fld = $myDataType.GetField('setupSource',[Reflection.BindingFlags]'Static,Public')
    $fld.SetValue($null, $source)

    $fldFilterStr = $myDataType.GetField(
      'filterString',
      [Reflection.BindingFlags] 'Static,Public'
    )
    $fldFilterStr.SetValue($null, '')

    # 4) Scan the media for all MSI/MSP packages
    $scanMedia = $myDataType.GetMethod('ScanSetupMedia',[Reflection.BindingFlags]'Static,Public')
    $scanMedia.Invoke($null, @())  # populates myData.sourcePkgs 

    # 5) Scan installed products & patches
    $scanProducts = $myDataType.GetMethod('ScanProducts',[Reflection.BindingFlags]'Static,Public')
    $scanProducts.Invoke($null, @())  # builds myData.rows with status OK/Mismatched/Missing

    # 6) Add any extra packages from LastUsedSource folders (this mirrors what AfterDone does under the covers)
    $addFromLast = $myDataType.GetMethod('AddMsiMspPackageFromLastUsedSource',[Reflection.BindingFlags]'Static,NonPublic')
    $addFromLast.Invoke($null, @())

    # 7) Generate the fix commands for each missing/mismatched row
    $updateFix = $myDataType.GetMethod('UpdateFixCommand',[Reflection.BindingFlags]'Static,Public')
    $updateFix.Invoke($null, @())  # populates each myRow.FixCommand 

    # 8) Pull out the resulting rows
    $rowsField = $myDataType.GetField('rows',[Reflection.BindingFlags]'Static,Public')
    $rows     = $rowsField.GetValue($null)  # a SortableBindingList<myRow>

    # 9) Filter to just missing or mismatched
    $badRows = $rows | Where-Object {$_.Status -in "Missing","Mismatched"}
 
    # 9.5) Populate FixCommand for files in our populated cache 
    foreach($row in $badRows | Where-Object {-Not $_.FixCommand}){
        if(Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)\$($row.PackageName)")){
            $row.FixCommand = "COPY `"\\$fileShareServer\FixMissingMSI\Cache\Products\$($row.ProductCode)\$($row.PackageCode)\$($row.PackageName)`" `"C:\Windows\Installer\$($row.CachedMsiMsp)`""
        }
        if(Test-Path("\\$fileShareServer\FixMissingMSI\Cache\Patches\$($row.ProductCode)\$($row.PatchCode)\$($row.PackageName)")){
            $row.FixCommand = "COPY `"\\$fileShareServer\FixMissingMSI\Cache\Patches\$($row.ProductCode)\$($row.PatchCode)\$($row.PackageName)`" `"C:\Windows\Installer\$($row.CachedMsiMsp)`""
        }
    }
     
    # update bad rows
    $badRowsWithFixCommand    = $badRows | Where-Object {$_.FixCommand}
    $badRowsWithoutFixCommand = $badRows | Where-Object {-Not $_.FixCommand} | Select-Object Status, PackageName, ProductName, Publisher, LastUsedSource, InstallSource, InstallDate, ProductCode, PackageCode, PatchCode, CachedMsiMsp, CachedMsiMspVersion, @{N="Hostname";E={$servername}}
    # export missing files data
    $badRowsWithoutFixCommand | Export-Csv -Path "\\$fileShareServer\FixMissingMSI\Reports\$servername.csv" -NoTypeInformation -Force

    # 10) Output findings + run fix commands:
    if($badRows.Count -gt 0){
        Write-Output "Missing: $(($badRows | Where-Object{$_.Status -eq "Missing"}).Count) Mismatched: $(($badRows | Where-Object{$_.Status -eq "Mismatched"}).Count) To be fixed: $($badRowsWithFixCommand.Count)"

        Foreach($row in $badRowsWithFixCommand){
            & cmd /c $row.FixCommand 
        }
    }
}

Stop-Transcript
