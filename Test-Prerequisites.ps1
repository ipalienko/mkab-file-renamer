[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Config not found: $ConfigPath. Copy config.example.json to config.json first."
}

$ConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$configDir = Split-Path -Parent $ConfigPath
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

Write-Host "[1/4] File share" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $config.Share.Root -PathType Container)) {
    throw "No access to file share: $($config.Share.Root)"
}
Write-Host "OK: $($config.Share.Root)" -ForegroundColor Green

Write-Host "[2/4] pdftotext" -ForegroundColor Cyan
$pdfToText = [string]$config.Tools.PdfToText
if (-not [System.IO.Path]::IsPathRooted($pdfToText)) {
    $pdfToText = Join-Path $configDir $pdfToText
}
if (-not (Test-Path -LiteralPath $pdfToText -PathType Leaf)) {
    throw "pdftotext.exe not found: $pdfToText"
}
& $pdfToText -v
if ($LASTEXITCODE -ne 0) { throw "pdftotext.exe returned exit code $LASTEXITCODE" }

Write-Host "[3/4] Microsoft Word COM" -ForegroundColor Cyan
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Write-Host "OK: Word $($word.Version)" -ForegroundColor Green
}
finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {}
    }
}

Write-Host "[4/4] SQL Server" -ForegroundColor Cyan
$cs = (
    "Data Source={0};Initial Catalog={1};User ID={2};Password={3};" +
    "Integrated Security=False;Encrypt={4};TrustServerCertificate={5};" +
    "Connect Timeout=15;Application Name=MKAB File Renamer Prerequisites;"
) -f $config.Sql.Server, $config.Sql.Database, $config.Sql.User, $config.Sql.Password, $config.Sql.Encrypt, $config.Sql.TrustServerCertificate

$conn = New-Object System.Data.SqlClient.SqlConnection($cs)
try {
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT TOP (1) MKABID, FAMILY, [NAME], OT, DATE_BD FROM dbo.hlt_MKAB;"
    [void]$cmd.ExecuteReader().Close()
    Write-Host "OK: SQL connection and dbo.hlt_MKAB read access" -ForegroundColor Green
}
finally {
    if ($null -ne $conn) {
        try { $conn.Close() } catch {}
        $conn.Dispose()
    }
}

Write-Host "All prerequisite checks passed." -ForegroundColor Green
