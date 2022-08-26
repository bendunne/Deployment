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
$DestPath = $env:APPDATA + "\Microsoft\AddIns\"

$registryPath2016 = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$Name = "OPEN3"
$value = "`"$XLAMFileName`""

# Removes XLAM file
If ((Test-Path ($DestPath + $XLAMFileName)) -eq $True){
   Remove-Item -Force -Path ($DestPath + $XLAMFileName)
}

# Removes Registry Value for Excel running 2016, 365

IF((Test-RegistryValue $registryPath2016 -Value $Name) -eq $True){
    Remove-ItemProperty -Path $registryPath2016 -Name $name -Force | Out-Null
  }

# Removes shortcuts created in AppData
Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\<Name of shortcut for copy script>.lnk" -Force
Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\<Name of shortcut for removal script>.lnk" -Force
