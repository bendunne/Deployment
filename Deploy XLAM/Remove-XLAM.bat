@ECHO OFF
pushd %~dp0
PowerShell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command "& '.\Remove-XLAM.ps1'"
popd