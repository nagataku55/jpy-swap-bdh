<#
.SYNOPSIS
    FastAPIサーバー(uvicorn main:app)をバックグラウンドで起動する。
    ログオン時にタスクスケジューラから呼ぶことを想定。

.DESCRIPTION
    - 指定ポートが既に使用中であれば、起動済みとみなして何もしない
      (ログオンのたびに重複起動してWinError 10048(アドレス使用中)に
      なるのを防ぐ。多重起動チェックがこのスクリプトの主な役割)
    - 未使用であれば、ウィンドウを表示せずuvicornを起動する
      (標準出力/エラーはProjectDir内のserver.log / server.err.logに残す)

.PARAMETER ProjectDir
    main.py が置かれているフォルダのフルパス。

.PARAMETER Port
    FastAPIサーバーのポート番号(既定: 8000)。ExportToDb.basのAPI_URL定数と
    一致させること。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File start_api_server.ps1 -ProjectDir "C:\python\swap"
#>
param(
    [string]$ProjectDir = "C:\python\swap",
    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"

$inUse = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($inUse) {
    Write-Output "port $Port は既に使用中です。起動済みとみなしてスキップします。"
    exit 0
}

$logPath = Join-Path $ProjectDir "server.log"
$errLogPath = Join-Path $ProjectDir "server.err.log"

Start-Process -FilePath "uvicorn" `
    -ArgumentList @("main:app", "--port", "$Port") `
    -WorkingDirectory $ProjectDir `
    -WindowStyle Hidden `
    -RedirectStandardOutput $logPath `
    -RedirectStandardError $errLogPath

Write-Output "uvicornをポート$Portで起動しました(ログ: $logPath)"

