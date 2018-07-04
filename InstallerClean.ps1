  
$Installer = New-Object -ComObject WindowsInstaller.Installer
$Type = $Installer.GetType()

function Get-MsiProducts {
	$Products = $Type.InvokeMember('Products', "GetProperty", $null, $Installer, $null)
	foreach ($Product In $Products) {
		$hash = @{}
		$hash.ProductCode = $Product
		$Attributes = @('Language', 'ProductName', 'PackageCode', 'Transforms', 'AssignmentType', 'PackageName', 'InstalledProductName', 'VersionString', 'RegCompany', 'RegOwner', 'ProductID', 'ProductIcon', 'InstallLocation', 'InstallSource', 'InstallDate', 'Publisher', 'LocalPackage', 'HelpLink', 'HelpTelephone', 'URLInfoAbout', 'URLUpdateInfo')		
		foreach ($Attribute In $Attributes) {
			$hash."$($Attribute)" = $null
		}
		foreach ($Attribute In $Attributes) {
			try {
				$hash."$($Attribute)" = $Type.InvokeMember('ProductInfo',"GetProperty", $null, $Installer, @($Product, $Attribute))
			} catch [System.Exception] {
			}
        }
        if($hash."LocalPackage"){
            if(test-path $hash."LocalPackage"){
                $hash.size=$(get-item $hash."LocalPackage").Length
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
	$Patches = $Type.InvokeMember('Patches',"GetProperty", $null, $Installer, @($product))
	foreach ($Patch In $Patches) {
		$hash = @{}
        $hash.ProductCode = $Product
        $hash.PatchCode=$Patch
		$Attributes = @('LocalPackage')		
		foreach ($Attribute In $Attributes) {
			$hash."$($Attribute)" = $null
		}
		foreach ($Attribute In $Attributes) {
			try {
				$hash."$($Attribute)" = $Type.InvokeMember('PatchInfo', 'GetProperty', $null, $Installer, @($Patch, $Attribute))
			} catch [System.Exception] {
				#$error[0]|format-list –force
			}
        }
        if($hash."LocalPackage"){
            if(test-path $hash."LocalPackage"){
                $hash.size=$(get-item $hash."LocalPackage").Length
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
        if(test-path $path){
            $path=get-item $path
            $extension=$path.Extension.ToLower()
            $DBOPENMODE=0
            $TABLENAME='Property'
            if($extension -eq '.msp'){
                $DBOPENMODE=32
                $TABLENAME="MsiPatchMetadata"
            }
            $msiProps = @{}
            $Database = $Type.InvokeMember("OpenDatabase", "InvokeMethod", $Null, $Installer, @($Path.FullName, $DBOPENMODE))
            $Query = "SELECT Property,Value FROM $TABLENAME"
            $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, ($Query))
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)|Out-Null
            $record=$view.gettype().invokemember("Fetch","InvokeMethod",$null,$view,$null)   
            # Loop thru the table
            while($record -ne $null) {
                $propName=$null
                $propValue=$null
                $propName=$record.gettype().invokeMember("StringData","GetProperty",$null,$record,1)
                $propValue= $record.gettype().invokeMember("StringData","GetProperty",$null,$record,2)
                $msiProps[$propName] =$propValue
                $record=$view.gettype().invokemember("Fetch","InvokeMethod",$null,$view,$null)
            }
            $view.gettype().invokemember("Close","InvokeMethod",$null,$view,$null)|Out-Null
            # Compose a unified object to express the MSI and MSP information
            # MSP  'DisplayName','ManufacturerName','Description', 'MoreInfoURL','TargetProductName'
            # MSI 'ProductName','Manufacturer','ProductVersion','ProductCode','UpgradeCode'
            if($extension -eq '.msi'){
                New-Object  -TypeName PSObject -Property @{
                    'DisplayName'=$msiProps['ProductName']
                    'Manufacturer'=$msiProps['Manufacturer']
                    'Version'=$msiProps['ProductVersion']
                    'PackageCode'=$msiProps['ProductCode']
                    'Description'=$msiProps['Description']
                    'TargetProductName'=$msiProps['TargetProductName']
                    'MoreInfoURL'=$msiProps['MoreInfoURL']
                    'Size'=$path.Length
                    'Path'=$path.FullName
                    'CreationTime'=$path.CreationTime
                }
            }elseif($extension -eq ".msp"){
                New-Object  -TypeName PSObject -Property @{
                    'DisplayName'=$msiProps['DisplayName']
                    'Manufacturer'=$msiProps['ManufacturerName']
                    'Version'=$msiProps['BuildNumber']
                    'PackageCode'=$msiProps['ProductCode']
                    'Description'=$msiProps['Description']
                    'TargetProductName'=$msiProps['TargetProductName']
                    'MoreInfoURL'=$msiProps['MoreInfoURL']
                    'Size'=$path.Length
                    'Path'=$path.FullName
                    'CreationTime'=$path.CreationTime
                }
            }       
        }     
	} catch {
		Write-Error $_.Exception.Message
	}
}

function filter_product{
    param(
        $productName
    )

    $PRODUCT_FILTER=@("adobe")

    $r=$PRODUCT_FILTER|?{$productName -like "*$_*"}
    if($r){
        return $true
    }else{
        return $false
    }
}


$products=Get-MsiProducts
$patches=$products|%{Get-MsiPatch -product $_.ProductCode}
$productsHash=@{}
$products|?{$_.LocalPackage}|%{$productsHash.add($_.LocalPackage,$true)}
$patchesHash=@{}
$patches|?{$_.LocalPackage}|%{if(!$patchesHash.ContainsKey($_.localPackage)){$patchesHash.add($_.LocalPackage,$true)}}
$InstallFolder="$($env:SystemRoot)\installer"
$files=dir  -Recurse -Include "*.msi","*.msp" -path $InstallFolder
$Files2=$files|%{
    if($productsHash.ContainsKey($_.FullName)){
        $_|Add-Member -MemberType NoteProperty -Name "installerState" -Value "InstalledProduct"          
    }elseif($patchesHash.ContainsKey($_.FullName)){
        $_|Add-Member -MemberType NoteProperty -Name "installerState" -Value "InstalledPatch"
    }else{
        $_|Add-Member -MemberType NoteProperty -Name "installerState" -Value "Orphaned"
    }
    $_
}

$groups=$files2|Group-Object -Property "installerState"
$groups|%{
    @{$($_.name)=($_.group|Measure-Object -Property Length -Sum).Sum}
}
$OrphanedFiles=$($groups|?{$_.name -eq 'Orphaned'}).Group
write-host "`n================================================================"

# issuses in powershell 2.0 
# 1. out-gridview has no -passthru switch 
# 2. there was no $psscriptroot env

if(!$PSScriptRoot){
    $psscriptroot=split-path $MyInvocation.MyCommand.path
}

function test-gridviewCmdlet{
    $result=$false
    (get-command out-gridview).ParameterSets[0].parameters|%{$_.name}|%{
        if($_ -eq 'passthru'){
            $result=$true
        }
    }
    return $result
}

if($OrphanedFiles){
    $ValidOrphanedFiles=($OrphanedFiles|%{
        $item=Get-MSIFileInfo -path $_.FullName;
        if((filter_product $item.DisplayName) -or (filter_product $item.Manufacturer)){
            # do nothing for this filtered products
        }else{
            $item
        }
    })

    if(test-gridviewCmdlet){
        write-host "And we will export Founded OrPhaned  Files to your desktop: "
        $selectedOrphanedFiles=$ValidOrphanedFiles|select DisplayName,Manufacturer,Size,Path,CreationTime|Out-GridView -PassThru -Title "select the Orphaned Files to delete"
        if($ValidOrphanedFiles){
            $ValidOrphanedFiles|Export-Csv -Path "$([environment]::GetFolderPath('desktop'))\ValidOrphanedFiles.$((get-date).ToString('yyyyMMddhhmmss')).csv" -NoClobber -NoTypeInformation -Encoding UTF8
        }
    
        if($selectedOrphanedFiles){
            $selectedOrphanedFiles|Export-Csv -Path "$([environment]::GetFolderPath('desktop'))\CleanedOrphanedFiles.$((get-date).ToString('yyyyMMddhhmmss')).csv" -NoClobber -NoTypeInformation -Encoding UTF8
            
            # delete code
            
            $selectedOrphanedFiles|remove-item -Force
        }
    }else{
        Write-host "We found your powershell does not support 'out-gridview' passthru switch" -ForegroundColor red
        write-host "So we just list the OrphanedFiles here" -ForegroundColor red

        $selectedOrphanedFiles=$ValidOrphanedFiles|select DisplayName,Manufacturer,Size,Path,CreationTime
        $selectedOrphanedFiles|Export-Csv -Path "$([environment]::GetFolderPath('desktop'))\ValidOrphanedFiles.$((get-date).ToString('yyyyMMddhhmmss')).csv" -NoClobber -NoTypeInformation -Encoding UTF8
        $selectedOrphanedFiles|Out-GridView -Title "select the Orphaned Files to delete"
        Start-Sleep -Seconds 30
    }
}else{
    write-host "No Orphaned Files Found! Exit Now" -ForegroundColor Green
}


