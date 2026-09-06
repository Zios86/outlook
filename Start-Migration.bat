@echo off
setlocal
chcp 65001 >nul

rem Папка назначения. Можно передать другой путь первым параметром BAT-файла.
set "DESTINATION=%~1"
if not defined DESTINATION set "DESTINATION=D:\OutlookArchive"

rem Для массового запуска средство управления должно запускать BAT от SYSTEM/администратора.
fltmc >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: запустите файл от имени администратора или SYSTEM.
    exit /b 1
)

rem Копируем скрипты в постоянную рабочую папку.
set "WORKDIR=%ProgramData%\OutlookPstMigration"
if not exist "%WORKDIR%" mkdir "%WORKDIR%"
if errorlevel 1 exit /b 2

copy /y "%~dp01-SystemStage.ps1" "%WORKDIR%\1-SystemStage.ps1" >nul
if errorlevel 1 exit /b 3
copy /y "%~dp02-UserStage.ps1" "%WORKDIR%\2-UserStage.ps1" >nul
if errorlevel 1 exit /b 4

rem Запускаем основной этап и возвращаем системе управления его код завершения.
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%WORKDIR%\1-SystemStage.ps1" -DestinationRoot "%DESTINATION%"
set "RESULT=%ERRORLEVEL%"

if not "%RESULT%"=="0" echo ОШИБКА: перенос завершился с кодом %RESULT%.
exit /b %RESULT%
