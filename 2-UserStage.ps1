#requires -Version 5.1
<#############################################################################
 ЭТАП 2. Автоматически запускается от имени сотрудника.
 Находит его PST, копирует, проверяет, переподключает Outlook и пишет результат.
#############################################################################>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$DestinationRoot)

$ErrorActionPreference = 'Stop'
$Sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$Root = Join-Path $env:ProgramData "OutlookPstMigrationSafeV7\$Sid"
$ResultFile = Join-Path $Root 'result.json'
$LogFile = Join-Path $Root 'migration.log'
$LockFile = Join-Path $Root 'running.lock'
$ProfilePath = $env:USERPROFILE
$UserKey = "$env:USERDOMAIN`_$env:USERNAME" -replace '[\\/:*?"<>|]', '_'
$Destination = Join-Path $DestinationRoot $UserKey
$TempFiles = [Collections.Generic.List[string]]::new()
$AddedStores = [Collections.Generic.List[string]]::new()
$RemovedStores = [Collections.Generic.List[string]]::new()
$Warnings = [Collections.Generic.List[string]]::new()
$Outlook = $null
$Mapi = $null
$ManifestFile = Join-Path $Root 'manifest.csv'

function Write-Log([string]$Text) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Text" | Add-Content -LiteralPath $LogFile -Encoding UTF8
}

function Write-Result(
    [string]$Status,
    [string]$Message,
    [int]$Count = 0,
    [long]$ProcessedBytes = 0,
    [long]$TotalBytes = 0
) {
    [pscustomobject]@{
        Version = 7; Status = $Status; Message = $Message; Count = $Count
        ProcessedBytes = $ProcessedBytes; TotalBytes = $TotalBytes
        Warnings = @($Warnings); LogFile = $LogFile; ManifestFile = $ManifestFile
    } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath "$ResultFile.tmp" -Encoding UTF8
    if (Test-Path -LiteralPath $ResultFile) {
        [IO.File]::Replace("$ResultFile.tmp", $ResultFile, $null)
    } else {
        Move-Item -LiteralPath "$ResultFile.tmp" -Destination $ResultFile
    }
}

function Get-FreeSpace([string]$Path) {
    if (-not ('DiskSpace.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace DiskSpace {
    public static class NativeMethods {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool GetDiskFreeSpaceEx(
            string directoryName,
            out ulong freeBytesAvailable,
            out ulong totalBytes,
            out ulong totalFreeBytes);
    }
}
'@
    }
    [UInt64]$Free = 0; [UInt64]$Total = 0; [UInt64]$TotalFree = 0
    if (-not [DiskSpace.NativeMethods]::GetDiskFreeSpaceEx($Path, [ref]$Free, [ref]$Total, [ref]$TotalFree)) {
        throw "Не удалось определить свободное место для $Path"
    }
    return [long]$Free
}

function Copy-PstFile([string]$Source, [string]$Target, [long]$TotalBytes, [ref]$OverallBytes) {
    $Buffer = New-Object byte[] (8MB)
    $InputStream = $null
    $OutputStream = $null
    $LastReport = [datetime]::MinValue
    try {
        $InputStream = [IO.File]::Open($Source, 'Open', 'Read', 'Read')
        $OutputStream = [IO.File]::Open($Target, 'CreateNew', 'Write', 'None')
        while (($Read = $InputStream.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
            $OutputStream.Write($Buffer, 0, $Read)
            $OverallBytes.Value += $Read
            if (((Get-Date) - $LastReport).TotalSeconds -ge 5) {
                $Percent = if ($TotalBytes -gt 0) { [math]::Min(100, [math]::Floor(100 * $OverallBytes.Value / $TotalBytes)) } else { 100 }
                $Message = "Копирование PST: $Percent%"
                Write-Log $Message
                Write-Result 'Running' $Message 0 $OverallBytes.Value $TotalBytes
                $LastReport = Get-Date
            }
        }
        $OutputStream.Flush()
    }
    finally {
        if ($OutputStream) { $OutputStream.Dispose() }
        if ($InputStream) { $InputStream.Dispose() }
    }
    (Get-Item -LiteralPath $Target).LastWriteTimeUtc = (Get-Item -LiteralPath $Source).LastWriteTimeUtc
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
        $Store = $null
        try {
            $Store = $Mapi.Stores.Item($i)
            if ($Store.FilePath -ieq $Path) { return $true }
        } finally {
            if ($Store) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Store) }
        }
    }
    return $false
}

function Remove-Store([string]$Path) {
    for ($i = $Mapi.Stores.Count; $i -ge 1; $i--) {
        $Store = $null; $RootFolder = $null
        try {
            $Store = $Mapi.Stores.Item($i)
            if ($Store.FilePath -ieq $Path) {
                $RootFolder = $Store.GetRootFolder()
                $Mapi.RemoveStore($RootFolder)
            }
        } finally {
            if ($RootFolder) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($RootFolder) }
            if ($Store) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Store) }
        }
    }
}

# Эксклюзивный lock защищает от одновременного запуска.
New-Item -Path $Root, $Destination -ItemType Directory -Force | Out-Null
$Lock = $null
try { $Lock = [IO.File]::Open($LockFile, 'OpenOrCreate', 'ReadWrite', 'None') }
catch { exit 75 }

try {
    Write-Result 'Running' 'Выполняется поиск PST.'
    Write-Log "Начало. Пользователь: $env:USERDOMAIN\$env:USERNAME"

    # Новый Outlook не поддерживает COM. Требуется установленный классический Outlook.
    if (-not [type]::GetTypeFromProgID('Outlook.Application')) {
        throw 'Классический Outlook не установлен. Новый Outlook не поддерживает подключение PST через COM.'
    }

    # Сначала получаем все PST, подключённые к Outlook. Они могут лежать на любом диске.
    $Outlook = New-Object -ComObject Outlook.Application
    $Mapi = $Outlook.GetNamespace('MAPI')
    $Connected = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    for ($i = 1; $i -le $Mapi.Stores.Count; $i++) {
        $Store = $null
        try {
            $Store = $Mapi.Stores.Item($i)
            $Path = $Store.FilePath
            if ($Path -and [IO.Path]::GetExtension($Path) -ieq '.pst' -and (Test-Path -LiteralPath $Path)) {
                if ($Path -notlike "$Destination\*") { [void]$Connected.Add($Path) }
            }
        } finally {
            if ($Store) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Store) }
        }
    }

    # Дополнительно ищем неподключённые PST внутри профиля пользователя.
    $Sources = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($Path in $Connected) { [void]$Sources.Add($Path) }
    $SearchErrors = @()
    Get-ChildItem -LiteralPath $ProfilePath -Filter '*.pst' -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +SearchErrors |
        Where-Object FullName -notlike "$Destination\*" |
        ForEach-Object { [void]$Sources.Add($_.FullName) }
    if ($SearchErrors.Count -gt 0) {
        [void]$Warnings.Add("Некоторые папки профиля недоступны для поиска: $($SearchErrors.Count)")
        Write-Log $Warnings[$Warnings.Count - 1]
    }

    if ($Sources.Count -eq 0) {
        Write-Log 'PST не найдены.'
        Write-Result 'Success' 'PST не найдены.' 0
        return
    }
    Write-Log "Найдено PST: $($Sources.Count)"

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

    # Проверяем объём данных и свободное место до начала копирования.
    $BytesToCopy = [long](($Plan | Where-Object { -not (Test-Path -LiteralPath $_.NewPath) } |
        ForEach-Object { (Get-Item -LiteralPath $_.OldPath).Length } | Measure-Object -Sum).Sum)
    $FreeBytes = Get-FreeSpace $Destination
    $ReserveBytes = 512MB
    if ($FreeBytes -lt ($BytesToCopy + $ReserveBytes)) {
        $NeedGb = [math]::Ceiling(($BytesToCopy + $ReserveBytes) / 1GB)
        $FreeGb = [math]::Round($FreeBytes / 1GB, 2)
        throw "Недостаточно места. Требуется около $NeedGb ГБ, свободно $FreeGb ГБ."
    }
    Write-Log "К копированию: $([math]::Round($BytesToCopy / 1GB, 2)) ГБ; свободно: $([math]::Round($FreeBytes / 1GB, 2)) ГБ."
    $CopiedBytes = [long]0

    # Просим Outlook закрыться штатно. Принудительное завершение запрещено.
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Mapi); $Mapi = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook); $Outlook = $null
    $OutlookProcesses = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)
    foreach ($Process in $OutlookProcesses) { [void]$Process.CloseMainWindow() }
    if ($OutlookProcesses.Count -gt 0) {
        Wait-Process -Id $OutlookProcesses.Id -Timeout 30 -ErrorAction SilentlyContinue
    }
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
        throw 'Outlook не закрылся за 30 секунд. Сохраните черновики, закройте Outlook и повторите запуск.'
    }

    # Сначала копируем и проверяем абсолютно все файлы. При ошибке будет выполнен откат.
    foreach ($Item in $Plan) {
        if (Test-Path -LiteralPath $Item.NewPath) {
            $SourceHash = (Get-FileHash -LiteralPath $Item.OldPath -Algorithm SHA256).Hash
            $DestinationHash = (Get-FileHash -LiteralPath $Item.NewPath -Algorithm SHA256).Hash
            if ($SourceHash -ne $DestinationHash) {
                throw "В месте назначения уже существует другой файл: $($Item.NewPath)"
            }
            $Item | Add-Member SourceHash $SourceHash -Force
            $Item | Add-Member DestinationHash $DestinationHash -Force
            Write-Log "Повторно использована проверенная копия: $($Item.NewPath)"
            continue
        }

        $Temp = "$($Item.NewPath).partial"
        # Остаток прерванного предыдущего запуска не должен мешать новой попытке.
        if (Test-Path -LiteralPath $Temp) {
            Remove-Item -LiteralPath $Temp -Force
            Write-Log "Удалён незавершённый временный файл предыдущего запуска: $Temp"
        }
        [void]$TempFiles.Add($Temp)
        Copy-PstFile -Source $Item.OldPath -Target $Temp -TotalBytes $BytesToCopy -OverallBytes ([ref]$CopiedBytes)
        $SourceHash = (Get-FileHash -LiteralPath $Item.OldPath -Algorithm SHA256).Hash
        $DestinationHash = (Get-FileHash -LiteralPath $Temp -Algorithm SHA256).Hash
        if ($SourceHash -ne $DestinationHash) {
            Remove-Item -LiteralPath $Temp -Force -ErrorAction SilentlyContinue
            throw "Контрольная сумма не совпала: $($Item.OldPath)"
        }
        Move-Item -LiteralPath $Temp -Destination $Item.NewPath
        [void]$TempFiles.Remove($Temp)
        $Item | Add-Member SourceHash $SourceHash -Force
        $Item | Add-Member DestinationHash $DestinationHash -Force
        Write-Log "Скопирован: $($Item.OldPath) -> $($Item.NewPath)"
    }

    $Plan | Select-Object OldPath, NewPath, WasConnected, SourceHash, DestinationHash |
        Export-Csv -LiteralPath $ManifestFile -NoTypeInformation -Encoding UTF8

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

    # Возвращаем профиль Outlook только если уже меняли подключения.
    try {
        if ($AddedStores.Count -gt 0 -or $RemovedStores.Count -gt 0) {
            if (-not $Outlook) { $Outlook = New-Object -ComObject Outlook.Application }
            if (-not $Mapi) { $Mapi = $Outlook.GetNamespace('MAPI') }
            foreach ($Path in $RemovedStores) { if (-not (Test-Store $Path)) { $Mapi.AddStoreEx($Path, 2) } }
            foreach ($Path in $AddedStores) { if (Test-Store $Path) { Remove-Store $Path } }
        }
        # Удаляются только незавершённые .partial. Проверенные копии сохраняются.
        foreach ($Path in $TempFiles) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
        Write-Log 'Откат профиля выполнен. Исходники и проверенные копии сохранены.'
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
    Remove-Item -LiteralPath "$ResultFile.tmp" -Force -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
