@echo off
cd %~dp0
Start /min powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ".\Install_&_Update.ps1"
