@echo off
cd %~dp0
Start /min powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ".\CCS-Tools-Launcher.ps1"