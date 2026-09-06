@echo off
setlocal EnableExtensions
title Outlook PST Migration SAFE V6

echo ================================================
echo  SAFE V6 - SOURCE PST FILES WILL NOT BE DELETED
echo ================================================
echo.
echo The BAT file started successfully.
echo Press any key to continue.
pause >nul

set "DESTINATION=C:\OutlookArchive"
if not "%~1"=="" set "DESTINATION=%~1"
set "WORKDIR=%ProgramData%\OutlookPstMigrationSafeV6"
set "LAUNCHLOG=%TEMP%\OutlookPstMigration-SAFE-V6.log"
set "RESULT=1"

echo [%date% %time%] Start SAFE V6 >"%LAUNCHLOG%"
echo Destination: %DESTINATION% >>"%LAUNCHLOG%"

fltmc >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Administrator rights are required.
    echo Right-click this BAT file and select Run as administrator.
    echo Administrator rights are missing. >>"%LAUNCHLOG%"
    goto finish
)

if not exist "%~dp01-SystemStage.ps1" (
    echo.
    echo ERROR: 1-SystemStage.ps1 was not found next to this BAT file.
    echo Extract the entire ZIP archive before starting.
    echo 1-SystemStage.ps1 was not found. >>"%LAUNCHLOG%"
    goto finish
)
if not exist "%~dp02-UserStage.ps1" (
    echo.
    echo ERROR: 2-UserStage.ps1 was not found next to this BAT file.
    echo Extract the entire ZIP archive before starting.
    echo 2-UserStage.ps1 was not found. >>"%LAUNCHLOG%"
    goto finish
)

if not exist "%WORKDIR%" mkdir "%WORKDIR%"
if errorlevel 1 (
    echo ERROR: Cannot create the working directory.
    echo Cannot create %WORKDIR%. >>"%LAUNCHLOG%"
    goto finish
)

copy /y "%~dp01-SystemStage.ps1" "%WORKDIR%\1-SystemStage.ps1" >>"%LAUNCHLOG%" 2>&1
if errorlevel 1 goto copyerror
copy /y "%~dp02-UserStage.ps1" "%WORKDIR%\2-UserStage.ps1" >>"%LAUNCHLOG%" 2>&1
if errorlevel 1 goto copyerror

echo.
echo Searching and safely copying PST files...
echo Source PST files will not be deleted.
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%WORKDIR%\1-SystemStage.ps1" -DestinationRoot "%DESTINATION%" >>"%LAUNCHLOG%" 2>&1
set "RESULT=%ERRORLEVEL%"

echo.
type "%LAUNCHLOG%"
echo.
if "%RESULT%"=="0" (
    echo DONE. Source PST files were preserved.
) else (
    echo ERROR. Migration was not completed. Source PST files should be preserved.
)
goto finish

:copyerror
echo.
echo ERROR: Cannot copy the PowerShell scripts.
echo Cannot copy the PowerShell scripts. >>"%LAUNCHLOG%"

:finish
echo.
echo Launcher log: "%LAUNCHLOG%"
echo Press any key to close this window.
pause >nul
exit /b %RESULT%
