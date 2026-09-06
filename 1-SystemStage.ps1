#requires -Version 5.1
<#############################################################################
 ЭТАП 1. Запускается от администратора или SYSTEM.
 Определяет сотрудника, создаёт пользовательское задание, ждёт результат,
 записывает итог и удаляет задание из Планировщика.
#############################################################################>
[CmdletBinding()]
param(
    [string]$DestinationRoot = 'C:\OutlookArchive',
    [int]$TimeoutMinutes = 120
)

$ErrorActionPreference = 'Stop'
$Root = Join-Path $env:ProgramData 'OutlookPstMigrationSafeV7'
$UserScript = Join-Path $Root '2-UserStage.ps1'
if (-not (Test-Path -LiteralPath $UserScript)) { throw "Не найден $UserScript" }

# Сначала берём пользователя активной консоли. Explorer используется только как резерв.
$User = (Get-CimInstance Win32_ComputerSystem).UserName
if (-not $User) {
    $Users = @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" | ForEach-Object {
        $Owner = Invoke-CimMethod -InputObject $_ -MethodName GetOwner
        if ($Owner.ReturnValue -eq 0 -and $Owner.User) { "$($Owner.Domain)\$($Owner.User)" }
    } | Sort-Object -Unique)
    if ($Users.Count -eq 1) { $User = $Users[0] }
}
if (-not $User) { throw 'Активный пользователь не найден.' }

# Получаем SID и реальный путь профиля пользователя.
$Sid = ([Security.Principal.NTAccount]$User).Translate([Security.Principal.SecurityIdentifier]).Value
$Profile = Get-CimInstance Win32_UserProfile | Where-Object SID -eq $Sid | Select-Object -First 1
if (-not $Profile) { throw "Профиль $User не найден." }

# Проверяем место назначения.
$Drive = Split-Path $DestinationRoot -Qualifier
if ($Drive -and -not (Test-Path -LiteralPath $Drive)) { throw "Диск $Drive отсутствует." }
$UserFolder = $User -replace '[\\/:*?"<>|]', '_'
$Destination = Join-Path $DestinationRoot $UserFolder
$UserWork = Join-Path $Root $Sid
$ResultFile = Join-Path $UserWork 'result.json'
$LockFile = Join-Path $UserWork 'running.lock'
$ControllerLockFile = Join-Path $UserWork 'controller.lock'
$TaskName = "Outlook PST safe migration v7 - $Sid"

# Предварительная проверка прав. Планировщик задач и ACL требуют администратора/SYSTEM.
$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$PrincipalCheck = New-Object Security.Principal.WindowsPrincipal($CurrentIdentity)
if (-not $PrincipalCheck.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Для запуска требуются права администратора или SYSTEM.'
}

New-Item -Path $Root, $UserWork, $Destination -ItemType Directory -Force | Out-Null
& icacls.exe $UserWork /grant "${User}:(OI)(CI)M" /T /C | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Не удалось выдать права на $UserWork" }
& icacls.exe $Destination /grant "${User}:(OI)(CI)M" /T /C | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Не удалось выдать права на $Destination" }

# Атомарно блокируем одновременный запуск двух системных этапов.
$ControllerLock = $null
try { $ControllerLock = [IO.File]::Open($ControllerLockFile, 'OpenOrCreate', 'ReadWrite', 'None') }
catch { throw 'Другой экземпляр переноса для этого пользователя уже запущен.' }

try {
    Remove-Item -LiteralPath $ResultFile -Force -ErrorAction SilentlyContinue

    # Запускаем второй этап именно в пользовательском сеансе.
    $Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$UserScript`" " +
                 "-DestinationRoot `"$DestinationRoot`""
    $Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $Arguments
    $Trigger = New-ScheduledTaskTrigger -AtLogOn -User $User
    $Principal = New-ScheduledTaskPrincipal -UserId $User -LogonType Interactive -RunLevel Limited
    $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes $TimeoutMinutes)

    $Result = $null
    $LastProgressMessage = $null
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
        -Principal $Principal -Settings $Settings -Force | Out-Null
    Start-ScheduledTask -TaskName $TaskName

    # Ждём подтверждение от второго этапа.
    $Deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        Start-Sleep -Seconds 2
        if (Test-Path -LiteralPath $ResultFile) {
            try { $Result = Get-Content $ResultFile -Raw | ConvertFrom-Json } catch { $Result = $null }
            if ($Result -and $Result.Message -and $Result.Message -ne $LastProgressMessage) {
                Write-Output $Result.Message
                $LastProgressMessage = $Result.Message
            }
            if ($Result -and $Result.Status -in @('Success', 'Failed')) { break }
        }
    } while ((Get-Date) -lt $Deadline)

    if (-not $Result -or $Result.Status -notin @('Success', 'Failed')) {
        Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        throw "Не получен результат за $TimeoutMinutes мин. Возможен PST с паролем. Исходники не удалены."
    }
    if ($Result.Status -eq 'Failed') { throw "$($Result.Message). Журнал: $($Result.LogFile)" }

    Write-Output "Успешно обработано PST: $($Result.Count). Исходники сохранены. Журнал: $($Result.LogFile)"
    # Машиночитаемая строка используется BAT-файлом для понятного итогового сообщения.
    Write-Output "PST_RESULT_COUNT=$($Result.Count)"
    if ($Result.Warnings) { Write-Warning ($Result.Warnings -join '; ') }
}
finally {
    # Задание больше не требуется и не останется запускаться при каждом входе.
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    if ($ControllerLock) { $ControllerLock.Dispose() }
    Remove-Item -LiteralPath $ControllerLockFile -Force -ErrorAction SilentlyContinue
}
