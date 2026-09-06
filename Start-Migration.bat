@echo off
setlocal EnableExtensions

set "DESTINATION=C:\OutlookArchive"
set "SILENT=0"
if /i "%~1"=="/silent" (
    set "SILENT=1"
) else if not "%~1"=="" (
    set "DESTINATION=%~1"
)
if /i "%~2"=="/silent" set "SILENT=1"

set "WORKDIR=%ProgramData%\OutlookPstMigrationSafeV8"
set "OUTPUT=%TEMP%\OutlookPstMigration-SAFE-V8.log"
set "RESULT=1"

fltmc >nul 2>&1
if errorlevel 1 goto finish
if not exist "%WORKDIR%" mkdir "%WORKDIR%"
if errorlevel 1 goto finish
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Test-Package.ps1" >nul 2>&1
if errorlevel 1 goto finish
copy /y "%~dp01-SystemStage.ps1" "%WORKDIR%\1-SystemStage.ps1" >nul
if errorlevel 1 goto finish
copy /y "%~dp02-UserStage.ps1" "%WORKDIR%\2-UserStage.ps1" >nul
if errorlevel 1 goto finish

powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%WORKDIR%\1-SystemStage.ps1" -DestinationRoot "%DESTINATION%" >"%OUTPUT%" 2>&1
set "RESULT=%ERRORLEVEL%"
type "%OUTPUT%"

:finish
if "%SILENT%"=="0" pause
exit /b %RESULT%
