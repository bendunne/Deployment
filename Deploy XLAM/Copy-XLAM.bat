@ECHO OFF
pushd %~dp0
PowerShell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command "& '.\Copy-XLAM.ps1'"
popd