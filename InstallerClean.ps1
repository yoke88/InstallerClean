
param(
    [switch] $quiet
)

$Installer = New-Object -ComObject WindowsInstaller.Installer
$Type = $Installer.GetType()



function check-outgridview {
    $x = get-command -Name Out-GridView -ErrorAction SilentlyContinue
    if (!$x) {
        return $false
    }
    else {
        if ($x.Parameters.ContainsKey("PassThru")) {
            return $true
        }
        else {
            return $false
        }
    }
}
function Get-MsiProducts {
    $Products = $Type.InvokeMember('Products', "GetProperty", $null, $Installer, $null)
    foreach ($Product In $Products) {
        $hash = @{ }
        $hash.ProductCode = $Product
        $Attributes = @('Language', 'ProductName', 'PackageCode', 'Transforms', 'AssignmentType', 'PackageName', 'InstalledProductName', 'VersionString', 'RegCompany', 'RegOwner', 'ProductID', 'ProductIcon', 'InstallLocation', 'InstallSource', 'InstallDate', 'Publisher', 'LocalPackage', 'HelpLink', 'HelpTelephone', 'URLInfoAbout', 'URLUpdateInfo')		
        foreach ($Attribute In $Attributes) {
            $hash."$($Attribute)" = $null
        }
        foreach ($Attribute In $Attributes) {
            try {
                $hash."$($Attribute)" = $Type.InvokeMember('ProductInfo', "GetProperty", $null, $Installer, @($Product, $Attribute))
            }
            catch [System.Exception] {
            }
        }
        if ($hash."LocalPackage") {
            if (test-path $hash."LocalPackage") {
                $hash.size = $(get-item $hash."LocalPackage").Length
            }
        }
        New-Object -TypeName PSObject -Property $hash
    }
}

function Get-MsiPatch {
    [cmdletbinding()]
    param(
        $product
    )
    $Patches = $Type.InvokeMember('Patches', "GetProperty", $null, $Installer, @($product))
    foreach ($Patch In $Patches) {
        $hash = @{ }
        $hash.ProductCode = $Product
        $hash.PatchCode = $Patch
        $Attributes = @('LocalPackage')		
        foreach ($Attribute In $Attributes) {
            $hash."$($Attribute)" = $null
        }
        foreach ($Attribute In $Attributes) {
            try {
                $hash."$($Attribute)" = $Type.InvokeMember('PatchInfo', 'GetProperty', $null, $Installer, @($Patch, $Attribute))
            }
            catch [System.Exception] {
                #$error[0]|format-list –force
            }
        }
        if ($hash."LocalPackage") {
            if (test-path $hash."LocalPackage") {
                $hash.size = $(get-item $hash."LocalPackage").Length
            }
        }
        New-Object -TypeName PSObject -Property $hash
    }
}

function Get-MSIFileInfo {
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory = $true)]$Path		
    )	
    try {
        if (test-path $path) {
            $path = get-item $path
            $extension = $path.Extension.ToLower()
            $DBOPENMODE = 0
            $TABLENAME = 'Property'
            if ($extension -eq '.msp') {
                $DBOPENMODE = 32
                $TABLENAME = "MsiPatchMetadata"
            }
            $msiProps = @{ }
            $Database = $Type.InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @($Path.FullName, $DBOPENMODE))
            $Query = "SELECT Property,Value FROM $TABLENAME"
            $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, ($Query))
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) | Out-Null
            $record = $view.gettype().invokemember("Fetch", "InvokeMethod", $null, $view, $null)   
            # Loop thru the table
            while ($Null -ne $record) {
                $propName = $null
                $propValue = $null
                $propName = $record.gettype().invokeMember("StringData", "GetProperty", $null, $record, 1)
                $propValue = $record.gettype().invokeMember("StringData", "GetProperty", $null, $record, 2)
                $msiProps[$propName] = $propValue
                $record = $view.gettype().invokemember("Fetch", "InvokeMethod", $null, $view, $null)
            }
            $view.gettype().invokemember("Close", "InvokeMethod", $null, $view, $null) | Out-Null
            # Compose a unified object to express the MSI and MSP information
            # MSP  'DisplayName','ManufacturerName','Description', 'MoreInfoURL','TargetProductName'
            # MSI 'ProductName','Manufacturer','ProductVersion','ProductCode','UpgradeCode'
            if ($extension -eq '.msi') {
                New-Object  -TypeName PSObject -Property @{
                    'DisplayName'       = $msiProps['ProductName']
                    'Manufacturer'      = $msiProps['Manufacturer']
                    'Version'           = $msiProps['ProductVersion']
                    'PackageCode'       = $msiProps['ProductCode']
                    'Description'       = $msiProps['Description']
                    'TargetProductName' = $msiProps['TargetProductName']
                    'MoreInfoURL'       = $msiProps['MoreInfoURL']
                    'Size'              = $path.Length
                    'Path'              = $path.FullName
                    'CreationTime'      = $path.CreationTime
                }
            }
            elseif ($extension -eq ".msp") {
                New-Object  -TypeName PSObject -Property @{
                    'DisplayName'       = $msiProps['DisplayName']
                    'Manufacturer'      = $msiProps['ManufacturerName']
                    'Version'           = $msiProps['BuildNumber']
                    'PackageCode'       = $msiProps['ProductCode']
                    'Description'       = $msiProps['Description']
                    'TargetProductName' = $msiProps['TargetProductName']
                    'MoreInfoURL'       = $msiProps['MoreInfoURL']
                    'Size'              = $path.Length
                    'Path'              = $path.FullName
                    'CreationTime'      = $path.CreationTime
                }
            }       
        }     
    }
    catch {
        Write-Error $_.Exception.Message
    }
}

function filter_product {
    param(
        $productName
    )

    $PRODUCT_FILTER = @("adobe")

    $r = $PRODUCT_FILTER | Where-Object { $productName -like "*$_*" }
    if ($r) {
        return $true
    }
    else {
        return $false
    }
}


$products = Get-MsiProducts
$patches = $products | ForEach-Object { Get-MsiPatch -product $_.ProductCode }
$productsHash = @{ }
$products | Where-Object { $_.LocalPackage } | ForEach-Object { $productsHash.add($_.LocalPackage, $true) }
$patchesHash = @{ }
$patches | Where-Object { $_.LocalPackage } | ForEach-Object { if (!$patchesHash.ContainsKey($_.localPackage)) { $patchesHash.add($_.LocalPackage, $true) } }
$InstallFolder = "$($env:SystemRoot)\installer"
$files = Get-ChildItem  -Recurse -Include "*.msi", "*.msp" -path $InstallFolder
$Files2 = $files | ForEach-Object {
    if ($productsHash.ContainsKey($_.FullName)) {
        $_ | Add-Member -MemberType NoteProperty -Name "installerState" -Value "InstalledProduct"          
    }
    elseif ($patchesHash.ContainsKey($_.FullName)) {
        $_ | Add-Member -MemberType NoteProperty -Name "installerState" -Value "InstalledPatch"
    }
    else {
        $_ | Add-Member -MemberType NoteProperty -Name "installerState" -Value "Orphaned"
    }
    $_
}

$groups = $files2 | Group-Object -Property "installerState"
$groups2 = @{ }
$groups | ForEach-Object {
    $key = $_.Name
    $size = ($_.group | Measure-Object -Property Length -Sum).Sum
    $groups2.Add($key, $size)
}
Write-Host "发现已安装产品文件尺寸有 $("{0} MB" -f ([math]::Floor($groups2['InstalledProduct']/1Mb)))"
Write-Host "发现有用的补丁文件尺寸有 $("{0} MB" -f ([math]::Floor($groups2['InstalledPatch']/1Mb)))"
Write-Host "发现可安全删除的补丁文件尺寸有 $("{0} MB" -f ([math]::Floor($groups2['Orphaned']/1Mb)))"
$OrphanedFiles = $($groups | Where-Object { $_.name -eq 'Orphaned' }).Group
if ($OrphanedFiles) {
    $ValidOrphanedFiles = ($OrphanedFiles | ForEach-Object {
            $item = Get-MSIFileInfo -path $_.FullName;
            if ((filter_product $item.DisplayName) -or (filter_product $item.Manufacturer)) {
                # do nothing for this filtered products
            }
            else {
                $item
            }
        })
    
    if (!$quiet -and (check-outgridview)) {
        $selectedOrphanedFiles = $ValidOrphanedFiles | Select-Object DisplayName, Manufacturer, Size, Path, CreationTime | Out-GridView -PassThru -Title "请选择要清理的补丁文件,此处列出的均可安全删除"
        if ($ValidOrphanedFiles) {
            $ValidOrphanedFiles | Export-Csv -Path $env:TEMP\ValidOrphanedFiles.$((get-date).ToString('yyyyMMddhhmmss')).csv -NoClobber -NoTypeInformation -Encoding UTF8
        }
        if ($selectedOrphanedFiles) {
            Write-Host "筛选后可删除的补丁文件尺寸有 $("{0} MB" -f ([math]::Floor(($selectedOrphanedFiles|Measure-Object -Property Size -Sum).sum/1Mb)))"
            $selectedOrphanedFiles | Export-Csv -Path $env:TEMP\CleanedOrphanedFiles.$((get-date).ToString('yyyyMMddhhmmss')).csv -NoClobber -NoTypeInformation -Encoding UTF8
            # delete code
            $selectedOrphanedFiles | remove-item -Force
        }
    }
    else {
        $selectedOrphanedFiles = $ValidOrphanedFiles| Where-Object {
            $_.Manufacturer -like '*Microsoft*'
        } 

        if ($selectedOrphanedFiles) {
            Write-Host "筛选后可删除的补丁文件尺寸有 $("{0} MB" -f ([math]::Floor(($selectedOrphanedFiles|Measure-Object -Property Size -Sum).sum/1Mb)))"
            $ExportCSVPath = "$($env:TEMP)\ValidOrphanedFiles.$((get-date).ToString('yyyyMMddhhmmss')).csv"
            Write-Host "筛选后的可以安全清除的补丁清单可在下面位置找到:"
            Write-Host $ExportCSVPath 
            $selectedOrphanedFiles | Export-Csv -Path $ExportCSVPath -NoClobber -NoTypeInformation -Encoding UTF8
        }
        if (!$quiet) {
            write-Host "请敲回车,表明您愿意继续删除过期补丁操作" -ForegroundColor Red
            write-host "如果您不同意继续,请按CTRL+C 或者关闭该窗口" -ForegroundColor Red   
            Read-Host 
        }
       
        $selectedOrphanedFiles | remove-item -Force

    }

}

write-host "清理完成,10秒后自动关闭窗口"
start-sleep -seconds 10