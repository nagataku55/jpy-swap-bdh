<#
.SYNOPSIS
    Excelをバックグラウンドで開き、DailyUpdate.RunDailyUpdate_Silent を実行する。
    Bloombergの引け値更新後(16:30頃)にWindowsタスクスケジューラから起動する想定。

.DESCRIPTION
    - Excelを非表示(Visible=$false)で起動し、指定ワークブックを開く
    - AutomationSecurityをmsoAutomationSecurityLow(1)にして、
      マクロ有効化の確認ダイアログを出さずにマクロを実行できるようにする(無人実行のため)
    - DailyUpdate.RunDailyUpdate_Silent を実行(再計算待ち→エラーチェック→保存→
      CSV監査出力(ExportHistLevels.bas)→FastAPI経由でSQLite取り込み(ExportToDb.bas)。
      結果はワークブック内のUpdateLogシートに記録される)
    - 処理後は必ずExcelを終了させる(COMオブジェクトが残りEXCEL.EXEが
      ゾンビ化しないよう、tryブロックの外でも確実にQuitする)

.PARAMETER WorkbookPath
    DailyUpdate.bas / ExportHistLevels.bas / ExportToDb.bas を組み込んだ
    マクロ有効ブック(.xlsm)のフルパス。

.NOTES
    - 事前に `uvicorn main:app` でFastAPIサーバーを起動しておくこと
      (タスクスケジューラの別タスクとして常駐起動するか、このスクリプトの前段で起動する)。
    - Bloomberg Terminalは対話ログインが前提のため、このタスクは
      「ユーザーがログオンしているときのみ実行する」設定でタスクスケジューラに登録すること
      (ユーザー未ログオン状態でのBloomberg接続は別製品(B-PIPE/Server API)が必要)。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File run_daily_update.ps1 -WorkbookPath "C:\python\swap\jpy_tona_data.xlsm"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    Write-Error "workbook not found: $WorkbookPath"
    exit 1
}
$WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path

$excel = $null
$workbook = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1   # msoAutomationSecurityLow: 確認ダイアログなしでマクロを実行する

    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $excel.Run("DailyUpdate.RunDailyUpdate_Silent")
    $workbook.Save()

    Write-Output "done: $WorkbookPath"
}
catch {
    Write-Error "run_daily_update failed: $($_.Exception.Message)"
    $exitCode = 1
}
finally {
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

exit $exitCode
