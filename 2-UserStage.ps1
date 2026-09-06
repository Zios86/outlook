#requires -Version 5.1
<#############################################################################
 SAFE V8 — пользовательский этап.
 Транзакционно копирует, проверяет и переподключает PST.
 Исходные PST никогда не удаляются.
#############################################################################>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$DestinationRoot)

$ErrorActionPreference = 'Stop'
$Sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$Root = Join-Path $env:ProgramData "OutlookPstMigrationSafeV8\$Sid"
$ResultFile = Join-Path $Root 'result.json'
$StateFile = Join-Path $Root 'state.json'
$LogFile = Join-Path $Root 'migration.log'
$ManifestFile = Join-Path $Root 'manifest.csv'
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
$State = $null
$CopiedBytes = [long]0
$BytesToCopy = [long]0

function Write-Log([string]$Text) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Text" | Add-Content -LiteralPath $LogFile -Encoding UTF8
}

function Write-AtomicJson([object]$Value, [string]$Path, [int]$Depth = 6) {
    $TempPath = "$Path.tmp"
    $Value | ConvertTo-Json -Depth $Depth | Set-Content -LiteralPath $TempPath -Encoding UTF8
    if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($TempPath, $Path, $null) }
    else { Move-Item -LiteralPath $TempPath -Destination $Path }
}

function Write-Result(
    [string]$Status, [string]$Message, [int]$Count = 0,
    [long]$ProcessedBytes = 0, [long]$TotalBytes = 0
) {
    Write-AtomicJson ([pscustomobject]@{
        Version = 8; Status = $Status; Message = $Message; Count = $Count
        ProcessedBytes = $ProcessedBytes; TotalBytes = $TotalBytes
        Warnings = @($Warnings); LogFile = $LogFile
        StateFile = $StateFile; ManifestFile = $ManifestFile
    }) $ResultFile 4
}

function Save-State([string]$OverallState) {
    if (-not $State) { return }
    $State.OverallState = $OverallState
    $State.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
    Write-AtomicJson $State $StateFile 8
}

function Get-FreeSpace([string]$Path) {
    if (-not ('DiskSpace.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace DiskSpace {
    public static class NativeMethods {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool GetDiskFreeSpaceEx(string directoryName,
            out ulong freeBytesAvailable, out ulong totalBytes, out ulong totalFreeBytes);
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

function Test-DestinationWrite([string]$Path) {
    $Probe = Join-Path $Path (".pst-write-test-{0}.tmp" -f [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($Probe, 'SAFE V8', [Text.Encoding]::UTF8)
        $Stream = [IO.File]::Open($Probe, 'Open', 'Read', 'Read')
        $Stream.Dispose()
    } finally { Remove-Item -LiteralPath $Probe -Force -ErrorAction SilentlyContinue }
}

function Get-ShortHash([string]$Text) {
    $Sha = [Security.Cryptography.SHA256]::Create()
    try {
        $Bytes = [Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant())
        return ([BitConverter]::ToString($Sha.ComputeHash($Bytes))).Replace('-', '').Substring(0, 12)
    } finally { $Sha.Dispose() }
}

function Get-SafeDestinationPath([IO.FileInfo]$File, [string]$SourcePath) {
    $Suffix = "_$(Get-ShortHash $SourcePath)$($File.Extension)"
    $MaximumNameLength = [math]::Min(180, 240 - $Destination.Length - 1)
    if ($MaximumNameLength -le $Suffix.Length) { throw "Слишком длинный путь назначения: $Destination" }
    $BaseLength = [math]::Min($File.BaseName.Length, $MaximumNameLength - $Suffix.Length)
    $SafeBase = $File.BaseName.Substring(0, $BaseLength) -replace '[\\/:*?"<>|]', '_'
    return Join-Path $Destination "$SafeBase$Suffix"
}

function Get-FileHashString([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Copy-PstFile([string]$Source, [string]$Target, [long]$TotalBytes, [ref]$OverallBytes) {
    $Buffer = New-Object byte[] (8MB)
    $Empty = New-Object byte[] 0
    $InputStream = $null; $OutputStream = $null
    $Sha = [Security.Cryptography.SHA256]::Create()
    $LastReport = [datetime]::MinValue
    try {
        # Запись источника блокируется на время копирования и проверки.
        $InputStream = [IO.File]::Open($Source, 'Open', 'Read', 'Read')
        $OutputStream = [IO.File]::Open($Target, 'CreateNew', 'Write', 'None')
        while (($Read = $InputStream.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
            $OutputStream.Write($Buffer, 0, $Read)
            [void]$Sha.TransformBlock($Buffer, 0, $Read, $Buffer, 0)
            $OverallBytes.Value += $Read
            if (((Get-Date) - $LastReport).TotalSeconds -ge 5) {
                $Percent = if ($TotalBytes -gt 0) { [math]::Min(100, [math]::Floor(100 * $OverallBytes.Value / $TotalBytes)) } else { 100 }
                $Message = "Копирование PST: $Percent%"
                Write-Log $Message
                Write-Result 'Running' $Message 0 $OverallBytes.Value $TotalBytes
                $LastReport = Get-Date
            }
        }
        [void]$Sha.TransformFinalBlock($Empty, 0, 0)
        $OutputStream.Flush($true)
        $OutputStream.Dispose(); $OutputStream = $null
        $SourceHash = ([BitConverter]::ToString($Sha.Hash)).Replace('-', '')
        $DestinationHash = Get-FileHashString $Target
        if ($SourceHash -ne $DestinationHash) { throw "Контрольная сумма не совпала: $Source" }
        (Get-Item -LiteralPath $Target).LastWriteTimeUtc = (Get-Item -LiteralPath $Source).LastWriteTimeUtc
        return [pscustomobject]@{ SourceHash = $SourceHash; DestinationHash = $DestinationHash }
    } finally {
        if ($OutputStream) { $OutputStream.Dispose() }
        if ($InputStream) { $InputStream.Dispose() }
        $Sha.Dispose()
    }
}

function Copy-PstWithRetry([object]$Item, [ref]$OverallBytes) {
    $Temp = "$($Item.NewPath).partial"
    [void]$TempFiles.Add($Temp)
    for ($Attempt = 1; $Attempt -le 3; $Attempt++) {
        $BytesBeforeAttempt = $OverallBytes.Value
        $Item.Attempts = $Attempt; $Item.Stage = 'Copying'; $Item.Error = $null
        Save-State 'Copying'
        Remove-Item -LiteralPath $Temp -Force -ErrorAction SilentlyContinue
        try {
            $Hashes = Copy-PstFile $Item.OldPath $Temp $BytesToCopy $OverallBytes
            Move-Item -LiteralPath $Temp -Destination $Item.NewPath
            [void]$TempFiles.Remove($Temp)
            $Item.SourceHash = $Hashes.SourceHash
            $Item.DestinationHash = $Hashes.DestinationHash
            $Item.Stage = 'Verified'; Save-State 'Copying'
            Write-Log "Скопирован и проверен: $($Item.OldPath) -> $($Item.NewPath)"
            return
        } catch {
            # Не учитываем байты неудачной попытки в общем прогрессе.
            $OverallBytes.Value = $BytesBeforeAttempt
            $Item.Error = $_.Exception.Message
            Remove-Item -LiteralPath $Temp -Force -ErrorAction SilentlyContinue
            Save-State 'Copying'
            if ($Attempt -eq 3) { throw }
            $Delay = [int][math]::Pow(2, $Attempt)
            Write-Log "Повтор копирования через $Delay сек. Причина: $($Item.Error)"
            Start-Sleep -Seconds $Delay
        }
    }
}

function Test-Store([string]$Path) {
    for ($i = 1; $i -le $Mapi.Stores.Count; $i++) {
        $Store = $null
        try {
            $Store = $Mapi.Stores.Item($i)
            if ($Store.FilePath -ieq $Path) { return $true }
        } finally { if ($Store) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Store) } }
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
                return
            }
        } finally {
            if ($RootFolder) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($RootFolder) }
            if ($Store) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Store) }
        }
    }
}

New-Item -Path $Root, $Destination -ItemType Directory -Force | Out-Null
$Lock = $null
try { $Lock = [IO.File]::Open($LockFile, 'OpenOrCreate', 'ReadWrite', 'None') }
catch { exit 75 }

try {
    Write-Result 'Running' 'Выполняется предварительная проверка.'
    Write-Log "Начало SAFE V8. Пользователь: $env:USERDOMAIN\$env:USERNAME"
    Test-DestinationWrite $Destination
    if (-not [type]::GetTypeFromProgID('Outlook.Application')) {
        throw 'Классический Outlook не установлен. Новый Outlook не поддерживает подключение PST через COM.'
    }

    $PreviousState = $null
    if (Test-Path -LiteralPath $StateFile) {
        try {
            $PreviousState = Get-Content -LiteralPath $StateFile -Raw | ConvertFrom-Json
            if ($PreviousState.Version -ne 8 -or $PreviousState.Destination -ine $Destination) { $PreviousState = $null }
        } catch {
            [void]$Warnings.Add('Предыдущее состояние повреждено и не использовано.')
            Write-Log $Warnings[$Warnings.Count - 1]
        }
    }

    $Outlook = New-Object -ComObject Outlook.Application
    $Mapi = $Outlook.GetNamespace('MAPI')
    $Connected = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    for ($i = 1; $i -le $Mapi.Stores.Count; $i++) {
        $Store = $null
        try {
            $Store = $Mapi.Stores.Item($i); $Path = $Store.FilePath
            if ($Path -and [IO.Path]::GetExtension($Path) -ieq '.pst' -and (Test-Path -LiteralPath $Path)) { [void]$Connected.Add($Path) }
        } finally { if ($Store) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Store) } }
    }

    $Sources = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($Path in $Connected) { if ($Path -notlike "$Destination\*") { [void]$Sources.Add($Path) } }
    $SearchErrors = @()
    Get-ChildItem -LiteralPath $ProfilePath -Filter '*.pst' -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable +SearchErrors |
        Where-Object { $_.FullName -notlike "$Destination\*" -and -not ($_.Attributes -band [IO.FileAttributes]::Offline) } |
        ForEach-Object { [void]$Sources.Add($_.FullName) }
    if ($PreviousState) {
        foreach ($OldItem in $PreviousState.Items) { if (Test-Path -LiteralPath $OldItem.OldPath) { [void]$Sources.Add($OldItem.OldPath) } }
    }
    if ($SearchErrors.Count -gt 0) {
        [void]$Warnings.Add("Некоторые папки профиля недоступны для поиска: $($SearchErrors.Count)")
        Write-Log $Warnings[$Warnings.Count - 1]
    }
    if ($Sources.Count -eq 0) {
        Write-Log 'PST не найдены.'; Write-Result 'Success' 'PST не найдены.' 0; return
    }

    $Items = [Collections.Generic.List[object]]::new()
    foreach ($OldPath in $Sources) {
        $File = Get-Item -LiteralPath $OldPath
        if ($File.Attributes -band [IO.FileAttributes]::Offline) { throw "PST доступен только как облачная заглушка: $OldPath" }
        $OldState = $null
        if ($PreviousState) { $OldState = $PreviousState.Items | Where-Object OldPath -ieq $OldPath | Select-Object -First 1 }
        $NewPath = if ($OldState -and $OldState.NewPath) { $OldState.NewPath } else { Get-SafeDestinationPath $File $OldPath }
        $WasConnected = $Connected.Contains($OldPath)
        if ($OldState -and $OldState.WasConnected) { $WasConnected = $true }
        [void]$Items.Add([pscustomobject]@{
            OldPath = $OldPath; NewPath = $NewPath; WasConnected = $WasConnected
            SourceLength = [long]$File.Length; SourceLastWriteUtc = $File.LastWriteTimeUtc.ToString('o')
            Stage = if ($OldState.Stage) { $OldState.Stage } else { 'Planned' }
            SourceHash = $OldState.SourceHash; DestinationHash = $OldState.DestinationHash
            Attempts = if ($OldState.Attempts) { [int]$OldState.Attempts } else { 0 }; Error = $null
        })
    }

    $State = [pscustomobject]@{
        Version = 8; RunId = [guid]::NewGuid().ToString('N')
        User = "$env:USERDOMAIN\$env:USERNAME"; Destination = $Destination
        CreatedAtUtc = if ($PreviousState.CreatedAtUtc) { $PreviousState.CreatedAtUtc } else { [datetime]::UtcNow.ToString('o') }
        UpdatedAtUtc = [datetime]::UtcNow.ToString('o'); OverallState = 'Planned'; Items = $Items
    }
    Save-State 'Planned'
    Write-Log "Найдено PST: $($Items.Count). План сохранён."

    $BytesToCopy = [long](($Items | Where-Object { -not (Test-Path -LiteralPath $_.NewPath) } | Measure-Object SourceLength -Sum).Sum)
    $FreeBytes = Get-FreeSpace $Destination; $ReserveBytes = 512MB
    if ($FreeBytes -lt ($BytesToCopy + $ReserveBytes)) {
        $NeedGb = [math]::Ceiling(($BytesToCopy + $ReserveBytes) / 1GB)
        $FreeGb = [math]::Round($FreeBytes / 1GB, 2)
        throw "Недостаточно места. Требуется около $NeedGb ГБ, свободно $FreeGb ГБ."
    }
    $DestinationRootPath = [IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Destination))
    if ($DestinationRootPath -and $DestinationRootPath -notlike '\\*') {
        try {
            $Volume = Get-Volume -DriveLetter $DestinationRootPath.Substring(0, 1) -ErrorAction Stop
            if ($Volume.FileSystem -eq 'FAT32' -and ($Items | Where-Object SourceLength -ge 4GB)) { throw 'FAT32 не поддерживает PST размером 4 ГБ и более.' }
        } catch {
            if ($_.Exception.Message -like 'FAT32*') { throw }
            [void]$Warnings.Add('Не удалось определить тип файловой системы назначения.')
        }
    }

    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Mapi); $Mapi = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook); $Outlook = $null
    $OutlookProcesses = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)
    foreach ($Process in $OutlookProcesses) { [void]$Process.CloseMainWindow() }
    if ($OutlookProcesses.Count -gt 0) { Wait-Process -Id $OutlookProcesses.Id -Timeout 30 -ErrorAction SilentlyContinue }
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Outlook не закрылся за 30 секунд. Сохраните черновики, закройте Outlook и повторите запуск.' }

    foreach ($Item in $Items) {
        if (Test-Path -LiteralPath $Item.NewPath) {
            $SourceHash = Get-FileHashString $Item.OldPath
            $DestinationHash = Get-FileHashString $Item.NewPath
            if ($SourceHash -ne $DestinationHash) { throw "В месте назначения уже существует другой файл: $($Item.NewPath)" }
            $Item.SourceHash = $SourceHash; $Item.DestinationHash = $DestinationHash; $Item.Stage = 'Verified'
            Save-State 'Copying'; Write-Log "Использована проверенная копия: $($Item.NewPath)"
        } else { Copy-PstWithRetry $Item ([ref]$CopiedBytes) }
    }

    $Items | Select-Object OldPath, NewPath, WasConnected, Stage, SourceLength, SourceHash, DestinationHash, Attempts |
        Export-Csv -LiteralPath $ManifestFile -NoTypeInformation -Encoding UTF8
    Save-State 'Verified'

    $Outlook = New-Object -ComObject Outlook.Application
    $Mapi = $Outlook.GetNamespace('MAPI')
    foreach ($Item in $Items | Where-Object WasConnected) {
        if (-not (Test-Path -LiteralPath $Item.NewPath)) { throw "Проверенная копия отсутствует: $($Item.NewPath)" }
        if ((Get-FileHashString $Item.NewPath) -ne $Item.DestinationHash) { throw "Копия изменилась перед подключением: $($Item.NewPath)" }
        if (-not (Test-Store $Item.NewPath)) {
            $Item.Stage = 'Attaching'; Save-State 'SwitchingOutlook'
            $Mapi.AddStoreEx($Item.NewPath, 2); [void]$AddedStores.Add($Item.NewPath)
        }
        if (-not (Test-Store $Item.NewPath)) { throw "Outlook не подключил $($Item.NewPath)" }
        $Item.Stage = 'Attached'; Save-State 'SwitchingOutlook'
    }
    foreach ($Item in $Items | Where-Object WasConnected) {
        if (Test-Store $Item.OldPath) {
            $Item.Stage = 'DetachingOld'; Save-State 'SwitchingOutlook'
            Remove-Store $Item.OldPath; [void]$RemovedStores.Add($Item.OldPath)
        }
        $Item.Stage = 'Completed'; Save-State 'SwitchingOutlook'
    }
    foreach ($Item in $Items | Where-Object { -not $_.WasConnected }) { $Item.Stage = 'Completed' }
    Save-State 'Completed'
    $Items | Select-Object OldPath, NewPath, WasConnected, Stage, SourceLength, SourceHash, DestinationHash, Attempts |
        Export-Csv -LiteralPath $ManifestFile -NoTypeInformation -Encoding UTF8
    Write-Log 'SAFE V8 завершён. Исходные PST сохранены.'
    Write-Result 'Success' 'Копирование и переподключение завершены. Исходные PST сохранены.' $Items.Count $BytesToCopy $BytesToCopy
}
catch {
    $MainError = $_.Exception.Message
    Write-Log "ОШИБКА: $MainError"
    try {
        if ($AddedStores.Count -gt 0 -or $RemovedStores.Count -gt 0) {
            if (-not $Outlook) { $Outlook = New-Object -ComObject Outlook.Application }
            if (-not $Mapi) { $Mapi = $Outlook.GetNamespace('MAPI') }
            foreach ($Path in $RemovedStores) { if (-not (Test-Store $Path)) { $Mapi.AddStoreEx($Path, 2) } }
            foreach ($Path in $AddedStores) { if (Test-Store $Path) { Remove-Store $Path } }
            if ($State) {
                foreach ($Item in $State.Items) { if ($Item.WasConnected) { $Item.Stage = 'Verified' } }
            }
        }
        foreach ($Path in $TempFiles) { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
        Save-State 'Failed'
        Write-Log 'Откат Outlook выполнен. Исходники и проверенные копии сохранены.'
    } catch {
        Write-Log "ОШИБКА ОТКАТА: $($_.Exception.Message). Исходные PST не удалялись."
        Save-State 'RecoveryRequired'
    }
    Write-Result 'Failed' $MainError 0 $CopiedBytes $BytesToCopy
}
finally {
    if ($Mapi) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Mapi) }
    if ($Outlook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) }
    if ($Lock) { $Lock.Dispose() }
    Remove-Item -LiteralPath $LockFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath "$ResultFile.tmp" -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath "$StateFile.tmp" -Force -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
