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
    if "%SILENT%"=="0" pause
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
    echo Подробный журнал находится в C:\ProgramData\OutlookPstMigration.
) else (
    findstr /c:"PST_RESULT_COUNT=0" "%OUTPUT%" >nul
    if errorlevel 1 (
        echo ГОТОВО: PST-архивы найдены и успешно перенесены в "%DESTINATION%".
    ) else (
        echo ГОТОВО: PST-архивы пользователя не найдены. Перенос не требуется.
    )
)

del /q "%OUTPUT%" >nul 2>&1
if "%SILENT%"=="0" pause
exit /b %RESULT%
