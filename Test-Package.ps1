#requires -Version 5.1
[CmdletBinding()]
param([string]$PackagePath = $PSScriptRoot)

$ErrorActionPreference = 'Stop'
$Failed = $false
$PowerShellFiles = @('1-SystemStage.ps1', '2-UserStage.ps1', 'Test-Package.ps1')
$BatchFiles = @('Start-Local-SAFE.bat', 'Start-Migration.bat')

foreach ($Name in $PowerShellFiles) {
    $Path = Join-Path $PackagePath $Name
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "Missing file: $Name" -ErrorAction Continue
        $Failed = $true
        continue
    }

    $Bytes = [IO.File]::ReadAllBytes($Path)
    if ($Bytes.Length -lt 3 -or $Bytes[0] -ne 0xEF -or $Bytes[1] -ne 0xBB -or $Bytes[2] -ne 0xBF) {
        Write-Error "UTF-8 BOM is missing: $Name" -ErrorAction Continue
        $Failed = $true
    }

    $Tokens = $null; $Errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$Tokens, [ref]$Errors)
    if ($Errors.Count -gt 0) {
        foreach ($ParseError in $Errors) { Write-Error "${Name}: $($ParseError.Message)" -ErrorAction Continue }
        $Failed = $true
    }
}

foreach ($Name in $BatchFiles) {
    $Path = Join-Path $PackagePath $Name
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "Missing file: $Name" -ErrorAction Continue
        $Failed = $true
        continue
    }
    $Bytes = [IO.File]::ReadAllBytes($Path)
    if ($Bytes | Where-Object { $_ -gt 0x7F }) {
        Write-Error "BAT is not ASCII: $Name" -ErrorAction Continue
        $Failed = $true
    }
}

$UserScript = Get-Content -LiteralPath (Join-Path $PackagePath '2-UserStage.ps1') -Raw
if ($UserScript -match '(Remove-Item|Move-Item|\[IO\.File\]::Delete)[^\r\n]*\$Item\.OldPath') {
    Write-Error 'Unsafe source PST deletion command detected.' -ErrorAction Continue
    $Failed = $true
}

$RuntimeScripts = @('1-SystemStage.ps1', '2-UserStage.ps1')
$AllScripts = Get-Content -LiteralPath ($RuntimeScripts | ForEach-Object { Join-Path $PackagePath $_ }) -Raw
if ($AllScripts -match 'Stop-Process[^\r\n]*OUTLOOK|Get-Process[^\r\n]*OUTLOOK[^\r\n]*\|[^\r\n]*Stop-Process') {
    Write-Error 'Forced Outlook termination command detected.' -ErrorAction Continue
    $Failed = $true
}
if ($AllScripts -match 'OutlookPstMigrationSafeV[0-7]') {
    Write-Error 'Reference to an obsolete work directory detected.' -ErrorAction Continue
    $Failed = $true
}

$RequiredV8Markers = @(
    "StateFile = Join-Path `$Root 'state.json'",
    "Save-State 'Planned'",
    "Save-State 'Verified'",
    "Save-State 'Completed'",
    'TransformFinalBlock',
    'Flush($true)'
)
foreach ($Marker in $RequiredV8Markers) {
    if ($UserScript.IndexOf($Marker, [StringComparison]::Ordinal) -lt 0) {
        Write-Error "Required SAFE V8 marker is missing: $Marker" -ErrorAction Continue
        $Failed = $true
    }
}

$SystemScript = Get-Content -LiteralPath (Join-Path $PackagePath '1-SystemStage.ps1') -Raw
if ($SystemScript -match 'New-ScheduledTaskTrigger\s+-AtLogOn') {
    Write-Error 'Persistent logon trigger detected.' -ErrorAction Continue
    $Failed = $true
}

if ($Failed) { exit 1 }
Write-Output 'PACKAGE_TEST=SUCCESS'
exit 0
