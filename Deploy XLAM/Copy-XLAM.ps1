function Test-RegistryValue {

param (

 [parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Path,

[parameter(Mandatory=$true)]
 [ValidateNotNullOrEmpty()]$Value
)

try {

Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
 return $true
 }

catch {

return $false

}

}

# Variables
$XLAMFileName = "<Add-in file name and extension>"
$sourcePath = "<Add-in source folder path>" + $XLAMFileName
$DestPath = $env:APPDATA + "\Microsoft\AddIns\"

$registryPath2016 = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$Name = "OPEN3"
$value = "`"$XLAMFileName`""


# If the XLAM doesn't exist then copy
If ((Test-Path ($DestPath + $XLAMFileName)) -eq $false){
    Copy-Item -Path $sourcePath -Destination $DestPath
}
# Else checks hash of the file from source and destination and compares if the same it won't copy otherwise it will
else{
    # check hashes
    $ref_file_hash_l = (Get-FileHash $sourcePath -ErrorAction SilentlyContinue).Hash
    $ref_file_hash_r = (Get-FileHash ($DestPath + $XLAMFileName)).Hash

    # compare
    if ($ref_file_hash_l -eq $ref_file_hash_r) {
        Write-Output "Hashes match: No update necessary"
    }
    # update required
    else {
        Write-Output "Hash mismatch: Updating local copy"

           Copy-Item -Force -Path $sourcePath -Destination $DestPath
        }
}

# Adds registry key to tell Excel to open the add-in for Excel running 2016, 365
IF((Test-RegistryValue $registryPath2016 -Value $Name) -eq $false){
    New-ItemProperty -Path $registryPath2016 -Name $name -Value $value -PropertyType string -Force | Out-Null
  }
