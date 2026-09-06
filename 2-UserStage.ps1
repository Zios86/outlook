#requires -Version 5.1
<#############################################################################
 ЭТАП 2. Автоматически запускается от имени сотрудника.
 Находит его PST, копирует, проверяет, переподключает Outlook и пишет результат.
#############################################################################>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$DestinationRoot)

$ErrorActionPreference = 'Stop'
$Sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$Root = Join-Path $env:ProgramData "OutlookPstMigrationSafeV2\$Sid"
$ResultFile = Join-Path $Root 'result.json'
$LogFile = Join-Path $Root 'migration.log'
$LockFile = Join-Path $Root 'running.lock'
$ProfilePath = $env:USERPROFILE
$Destination = Join-Path $DestinationRoot (Split-Path $ProfilePath -Leaf)
$CreatedCopies = [Collections.Generic.List[string]]::new()
$TempFiles = [Collections.Generic.List[string]]::new()
$AddedStores = [Collections.Generic.List[string]]::new()
$RemovedStores = [Collections.Generic.List[string]]::new()
$Warnings = [Collections.Generic.List[string]]::new()
$Outlook = $null
$Mapi = $null

function Write-Log([string]$Text) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Text" | Add-Content -LiteralPath $LogFile -Encoding UTF8
}

function Write-Result([string]$Status, [string]$Message, [int]$Count = 0) {
    [pscustomobject]@{
        Status = $Status; Message = $Message; Count = $Count
        Warnings = @($Warnings); LogFile = $LogFile
    } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $ResultFile -Encoding UTF8
}

function Get-ShortHash([string]$Text) {
    $Sha = [Security.Cryptography.SHA256]::Create()
    try {
        $Bytes = [Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant())
        return ([BitConverter]::ToString($Sha.ComputeHash($Bytes))).Replace('-', '').Substring(0, 10)
    } finally { $Sha.Dispose() }
}

function Test-Store([string]$Path) {
    for ($i = 1; $i -le $Mapi.Stores.Count; $i++) {
        if ($Mapi.Stores.Item($i).FilePath -ieq $Path) { return $true }
    }
    return $false
}

function Remove-Store([string]$Path) {
    for ($i = $Mapi.Stores.Count; $i -ge 1; $i--) {
        $Store = $Mapi.Stores.Item($i)
        if ($Store.FilePath -ieq $Path) { $Mapi.RemoveStore($Store.GetRootFolder()) }
    }
}

# Эксклюзивный lock защищает от одновременного запуска.
New-Item -Path $Root, $Destination -ItemType Directory -Force | Out-Null
$Lock = $null
try { $Lock = [IO.File]::Open($LockFile, 'CreateNew', 'Write', 'None') }
catch { exit 0 }

try {
    Write-Result 'Running' 'Выполняется поиск PST.'
    Write-Log "Начало. Пользователь: $env:USERDOMAIN\$env:USERNAME"

    # Сначала получаем все PST, подключённые к Outlook. Они могут лежать на любом диске.
    $Outlook = New-Object -ComObject Outlook.Application
    $Mapi = $Outlook.GetNamespace('MAPI')
    $Connected = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    for ($i = 1; $i -le $Mapi.Stores.Count; $i++) {
        $Path = $Mapi.Stores.Item($i).FilePath
        if ($Path -and [IO.Path]::GetExtension($Path) -ieq '.pst' -and (Test-Path -LiteralPath $Path)) {
            if ($Path -notlike "$Destination\*") { [void]$Connected.Add($Path) }
        }
    }

    # Дополнительно ищем неподключённые PST внутри профиля пользователя.
    $Sources = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($Path in $Connected) { [void]$Sources.Add($Path) }
    Get-ChildItem -LiteralPath $ProfilePath -Filter '*.pst' -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object FullName -notlike "$Destination\*" |
        ForEach-Object { [void]$Sources.Add($_.FullName) }

    if ($Sources.Count -eq 0) {
        Write-Log 'PST не найдены.'
        Write-Result 'Success' 'PST не найдены.' 0
        return
    }
    Write-Log "Найдено PST: $($Sources.Count)"

    # Закрываем Outlook перед копированием открытых файлов.
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Mapi); $Mapi = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook); $Outlook = $null
    Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 3

    # Формируем постоянные уникальные имена по исходному пути. Повторный запуск не создаст дубликаты.
    $Plan = foreach ($OldPath in $Sources) {
        $File = Get-Item -LiteralPath $OldPath
        $NewName = "$($File.BaseName)_$(Get-ShortHash $OldPath)$($File.Extension)"
        [pscustomobject]@{
            OldPath = $OldPath
            NewPath = Join-Path $Destination $NewName
            WasConnected = $Connected.Contains($OldPath)
        }
    }

    # Сначала копируем и проверяем абсолютно все файлы. При ошибке будет выполнен откат.
    foreach ($Item in $Plan) {
        if (Test-Path -LiteralPath $Item.NewPath) {
            if ((Get-FileHash -LiteralPath $Item.OldPath).Hash -ne (Get-FileHash -LiteralPath $Item.NewPath).Hash) {
                throw "В месте назначения уже существует другой файл: $($Item.NewPath)"
            }
            Write-Log "Повторно использована проверенная копия: $($Item.NewPath)"
            continue
        }

        $Temp = "$($Item.NewPath).partial"
        [void]$TempFiles.Add($Temp)
        Copy-Item -LiteralPath $Item.OldPath -Destination $Temp -Force
        if ((Get-FileHash -LiteralPath $Item.OldPath).Hash -ne (Get-FileHash -LiteralPath $Temp).Hash) {
            Remove-Item -LiteralPath $Temp -Force -ErrorAction SilentlyContinue
            throw "Контрольная сумма не совпала: $($Item.OldPath)"
        }
        Move-Item -LiteralPath $Temp -Destination $Item.NewPath
        [void]$TempFiles.Remove($Temp)
        [void]$CreatedCopies.Add($Item.NewPath)
        Write-Log "Скопирован: $($Item.OldPath) -> $($Item.NewPath)"
    }

    # Подключаем новые копии только для архивов, ранее подключённых к Outlook.
    $Outlook = New-Object -ComObject Outlook.Application
    $Mapi = $Outlook.GetNamespace('MAPI')
    foreach ($Item in $Plan | Where-Object WasConnected) {
        if (-not (Test-Store $Item.NewPath)) {
            Write-Log "Подключение: $($Item.NewPath)"
            $Mapi.AddStoreEx($Item.NewPath, 2) # 2 = Unicode PST
            [void]$AddedStores.Add($Item.NewPath)
        }
        if (-not (Test-Store $Item.NewPath)) { throw "Outlook не подключил $($Item.NewPath)" }
    }

    # Только после подключения всех новых PST отключаем старые пути.
    foreach ($Item in $Plan | Where-Object WasConnected) {
        if (Test-Store $Item.OldPath) {
            Remove-Store $Item.OldPath
            [void]$RemovedStores.Add($Item.OldPath)
        }
    }

    # Безопасный режим: исходные PST не удаляем. После проверки администратор
    # сможет удалить их отдельно, когда убедится, что новые архивы работают.
    foreach ($Item in $Plan) {
        Write-Log "Исходник сохранён: $($Item.OldPath)"
    }

    Write-Log 'Копирование и переподключение завершены. Исходники сохранены.'
    Write-Result 'Success' 'Копирование завершено. Исходники сохранены.' $Plan.Count
}
catch {
    $MainError = $_.Exception.Message
    Write-Log "ОШИБКА: $MainError"

    # До удаления исходников возвращаем профиль Outlook в исходное состояние.
    try {
        if (-not $Outlook) { $Outlook = New-Object -ComObject Outlook.Application }
        if (-not $Mapi) { $Mapi = $Outlook.GetNamespace('MAPI') }
        foreach ($Path in $RemovedStores) { if (-not (Test-Store $Path)) { $Mapi.AddStoreEx($Path, 2) } }
        foreach ($Path in $AddedStores) { if (Test-Store $Path) { Remove-Store $Path } }
        foreach ($Path in $CreatedCopies) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
        foreach ($Path in $TempFiles) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
        Write-Log 'Откат выполнен, исходники сохранены.'
    } catch {
        Write-Log "ОШИБКА ОТКАТА: $($_.Exception.Message). Исходники не удалялись."
    }
    Write-Result 'Failed' $MainError
}
finally {
    if ($Mapi) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Mapi) }
    if ($Outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) }
    if ($Lock) { $Lock.Dispose() }
    Remove-Item -LiteralPath $LockFile -Force -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
