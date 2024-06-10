@echo off
echo Packaging application with PyInstaller...

pyinstaller --onefile --add-data "AD;AD" --add-data "Vmware;Vmware" --add-data "Update;Update" CCS-Tools-Launcher.py

if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller failed. Exiting.
    pause
    exit /b %ERRORLEVEL%
)

echo PyInstaller build complete.

set EXE_NAME=dist\CCS-Tools-Launcher.exe
set ZIP_NAME=CCS-Tools-Launcher.zip

if exist %EXE_NAME% (
    echo Zipping the executable...
    powershell -command "Compress-Archive -Path '%EXE_NAME%' -DestinationPath '%ZIP_NAME%'"

    if %ERRORLEVEL% EQU 0 (
        echo Zipping complete. The zip file is %ZIP_NAME%.
    ) else (
        echo Failed to zip the executable.
    )
) else (
    echo Executable not found.
)

pause
