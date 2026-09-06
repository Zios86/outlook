@echo off
setlocal
chcp 65001 >nul

rem По умолчанию архивы переносятся сюда.
set "DESTINATION=D:\OutlookArchive"
set "SILENT=0"

rem Другой путь можно передать первым параметром. /silent отключает паузу.
if /i "%~1"=="/silent" (
    set "SILENT=1"
) else if not "%~1"=="" (
    set "DESTINATION=%~1"
)
if /i "%~2"=="/silent" set "SILENT=1"

rem Для массового запуска средство управления должно запускать BAT от SYSTEM/администратора.
fltmc >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: запустите файл от имени администратора или SYSTEM.
    set "RESULT=1"
    goto :finish
)

rem Копируем скрипты в постоянную рабочую папку.
set "WORKDIR=%ProgramData%\OutlookPstMigrationSafeV2"
if not exist "%WORKDIR%" mkdir "%WORKDIR%"
if errorlevel 1 (
    echo ОШИБКА: не удалось создать "%WORKDIR%".
    set "RESULT=2"
    goto :finish
)

copy /y "%~dp01-SystemStage.ps1" "%WORKDIR%\1-SystemStage.ps1" >nul
if errorlevel 1 (
    echo ОШИБКА: рядом с BAT не найден файл 1-SystemStage.ps1.
    echo Сначала полностью распакуйте ZIP-архив.
    set "RESULT=3"
    goto :finish
)
copy /y "%~dp02-UserStage.ps1" "%WORKDIR%\2-UserStage.ps1" >nul
if errorlevel 1 (
    echo ОШИБКА: рядом с BAT не найден файл 2-UserStage.ps1.
    echo Сначала полностью распакуйте ZIP-архив.
    set "RESULT=4"
    goto :finish
)

echo.
echo Выполняется поиск PST-архивов пользователя...
echo Не закрывайте это окно. Outlook может быть автоматически закрыт.
echo.

rem Сохраняем вывод, чтобы показать понятный итог.
set "OUTPUT=%TEMP%\OutlookPstMigration-%RANDOM%.log"
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%WORKDIR%\1-SystemStage.ps1" -DestinationRoot "%DESTINATION%" >"%OUTPUT%" 2>&1
set "RESULT=%ERRORLEVEL%"

type "%OUTPUT%"
echo.

if not "%RESULT%"=="0" (
    echo ОШИБКА: перенос не выполнен. Код ошибки: %RESULT%.
    echo Подробный журнал находится в C:\ProgramData\OutlookPstMigrationSafeV2.
) else (
    findstr /c:"PST_RESULT_COUNT=0" "%OUTPUT%" >nul
    if errorlevel 1 (
        echo ГОТОВО: PST-архивы скопированы в "%DESTINATION%".
        echo Исходные PST-файлы сохранены и автоматически не удалялись.
    ) else (
        echo ГОТОВО: PST-архивы пользователя не найдены. Перенос не требуется.
    )
)

del /q "%OUTPUT%" >nul 2>&1

:finish
if not defined RESULT set "RESULT=1"
if "%SILENT%"=="0" pause
exit /b %RESULT%
