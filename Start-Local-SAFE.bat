@echo off
setlocal EnableExtensions
title Outlook PST Migration SAFE V2
chcp 65001 >nul

echo ================================================
echo  БЕЗОПАСНАЯ ВЕРСИЯ V2 - ИСХОДНЫЕ PST НЕ УДАЛЯЕТ
echo ================================================
echo.
echo Если это окно видно, BAT-файл запустился нормально.
echo Для продолжения нажмите любую клавишу.
pause >nul

set "DESTINATION=D:\OutlookArchive"
if not "%~1"=="" set "DESTINATION=%~1"
set "WORKDIR=%ProgramData%\OutlookPstMigrationSafeV2"
set "LAUNCHLOG=%TEMP%\OutlookPstMigration-SAFE-V2.log"
set "RESULT=1"

echo [%date% %time%] Запуск SAFE V2 >"%LAUNCHLOG%"
echo Папка назначения: %DESTINATION% >>"%LAUNCHLOG%"

fltmc >nul 2>&1
if errorlevel 1 (
    echo.
    echo ОШИБКА: нужны права администратора.
    echo Нажмите правой кнопкой на BAT и выберите «Запуск от имени администратора».
    echo Нет прав администратора. >>"%LAUNCHLOG%"
    goto :finish
)

if not exist "%~dp01-SystemStage.ps1" (
    echo.
    echo ОШИБКА: рядом нет файла 1-SystemStage.ps1.
    echo Полностью распакуйте ZIP, затем запускайте BAT из распакованной папки.
    echo Не найден 1-SystemStage.ps1. >>"%LAUNCHLOG%"
    goto :finish
)
if not exist "%~dp02-UserStage.ps1" (
    echo.
    echo ОШИБКА: рядом нет файла 2-UserStage.ps1.
    echo Полностью распакуйте ZIP, затем запускайте BAT из распакованной папки.
    echo Не найден 2-UserStage.ps1. >>"%LAUNCHLOG%"
    goto :finish
)

if not exist "%WORKDIR%" mkdir "%WORKDIR%"
if errorlevel 1 (
    echo ОШИБКА: не удалось создать рабочую папку.
    echo Ошибка создания %WORKDIR%. >>"%LAUNCHLOG%"
    goto :finish
)

copy /y "%~dp01-SystemStage.ps1" "%WORKDIR%\1-SystemStage.ps1" >>"%LAUNCHLOG%" 2>&1
if errorlevel 1 goto :copyerror
copy /y "%~dp02-UserStage.ps1" "%WORKDIR%\2-UserStage.ps1" >>"%LAUNCHLOG%" 2>&1
if errorlevel 1 goto :copyerror

echo.
echo Идёт поиск и безопасное копирование PST...
echo Исходные PST удаляться не будут.
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%WORKDIR%\1-SystemStage.ps1" -DestinationRoot "%DESTINATION%" >>"%LAUNCHLOG%" 2>&1
set "RESULT=%ERRORLEVEL%"

echo.
type "%LAUNCHLOG%"
echo.
if "%RESULT%"=="0" (
    echo ГОТОВО. Исходные PST сохранены.
) else (
    echo ОШИБКА. Перенос не завершён. Исходные PST не должны удаляться.
)
goto :finish

:copyerror
echo.
echo ОШИБКА: не удалось скопировать служебные скрипты.
echo Ошибка копирования файлов. >>"%LAUNCHLOG%"

:finish
echo.
echo Журнал запуска: "%LAUNCHLOG%"
echo Для закрытия окна нажмите любую клавишу.
pause >nul
exit /b %RESULT%
