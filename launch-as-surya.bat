@echo off
setlocal

set "RUNAS_USER=NEECHEWALA\surya"
set "APP_PATH=C:\Users\Anurag Singh\Desktop\JB Creations Labels 1.0.0.exe"

if not exist "%APP_PATH%" (
    echo App not found:
    echo %APP_PATH%
    pause
    exit /b 1
)

runas /savecred /user:%RUNAS_USER% "\"%APP_PATH%\""

if errorlevel 1 (
    echo Failed to launch %APP_PATH% as %RUNAS_USER%.
    pause
    exit /b %errorlevel%
)

endlocal