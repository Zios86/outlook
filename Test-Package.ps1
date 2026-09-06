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
if ($UserScript -match 'Remove-Item\s+-LiteralPath\s+\$Item\.OldPath') {
    Write-Error 'Unsafe source PST deletion command detected.' -ErrorAction Continue
    $Failed = $true
}

if ($Failed) { exit 1 }
Write-Output 'PACKAGE_TEST=SUCCESS'
exit 0
