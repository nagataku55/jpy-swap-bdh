Attribute VB_Name = "DailyUpdate"
Option Explicit

' BDH再計算後に、エラーが無ければ保存→CSV監査ログ出力(ExportHistLevels.bas)
' →FastAPI経由でSQLite取り込み(ExportToDb.bas)→config_fwd同期(ForwardSpec.bas)
' →フォワードスワップレートの欠損バックフィル(ForwardHistory.bas)まで行う。
'
' 呼び出し方は2通り:
'   - RunDailyUpdate_Manual  : シート上のボタンに割り当てる(結果をMsgBoxで表示)
'   - RunDailyUpdate_Silent  : run_daily_update.ps1(タスクスケジューラ)から呼ぶ(MsgBoxを出さない)
'
' エラーセルが1つでも見つかった場合は、その日は保存・CSV出力・DB取り込みを
' 一切行わない(中途半端なデータを残さないため)。結果はUpdateLogシートに追記する。
'
' ForwardSpec同期・ForwardHistory更新は、スポットデータの取り込み(上記)が成功した
' 後に行うが、ここで失敗してもその日のスポットデータ取り込み自体は既に成功している
' ので、処理全体を失敗扱いにはしない(ログに警告として残すのみ)。

Private Const DATA_START_CELL As String = "B3"      ' Date列の先頭データセル
Private Const LOG_SHEET As String = "UpdateLog"
Private Const RECALC_TIMEOUT_SEC As Long = 60

' データシート名はハードコードせず、Controlシートの設定セルから読む
Private Const CONTROL_SHEET As String = "Control"
Private Const DATA_SHEET_NAME_CELL As String = "B24"  ' 例: "JPY_TONA_Hist"


Public Sub RunDailyUpdate_Manual()
    RunDailyUpdate False
End Sub

Public Sub RunDailyUpdate_Silent()
    RunDailyUpdate True
End Sub

Private Sub RunDailyUpdate(silent As Boolean)
    LogStatus "更新処理を開始します"

    Dim dataSheet As String
    On Error GoTo ConfigError
    dataSheet = GetDataSheetName()
    On Error GoTo 0

    If Not WaitForRecalculation(RECALC_TIMEOUT_SEC) Then
        LogStatus "中止: 再計算が" & RECALC_TIMEOUT_SEC & "秒以内に完了しませんでした"
        NotifyIfInteractive silent, "再計算がタイムアウトしました。Bloombergの接続状況を確認してください。", vbExclamation
        Exit Sub
    End If

    Dim errCell As String
    errCell = FindErrorCell(dataSheet, DATA_START_CELL)
    If errCell <> "" Then
        LogStatus "中止: " & dataSheet & "!" & errCell & " にエラー値があります。本日分は取り込みません"
        NotifyIfInteractive silent, "エラーセルが見つかりました(" & dataSheet & "!" & errCell & ")。" & _
            vbCrLf & "本日分の取り込みは中止しました。Bloombergの接続を確認し、手動で再実行してください。", vbExclamation
        Exit Sub
    End If

    On Error GoTo SaveError
    ThisWorkbook.Save
    On Error GoTo 0

    Dim exportMsg As String
    exportMsg = ExportHistLevels.ExportAllSheetsToCsv_Quiet()
    LogStatus "CSVエクスポート: " & exportMsg

    If InStr(1, exportMsg, "エラー", vbTextCompare) = 1 Then
        LogStatus "中止: CSVエクスポートに失敗したためDB取り込みをスキップします"
        NotifyIfInteractive silent, "CSVエクスポートに失敗しました。" & vbCrLf & exportMsg, vbExclamation
        Exit Sub
    End If

    Dim dbMsg As String
    dbMsg = ExportToDb.RunExportToDb_Silent()
    LogStatus "DB取り込み: " & dbMsg

    If InStr(1, dbMsg, "エラー", vbTextCompare) = 1 Then
        LogStatus "中止: DB取り込みに失敗しました"
        NotifyIfInteractive silent, "DB取り込みに失敗しました。UpdateLogシートとFastAPIサーバーのログを確認してください。", vbExclamation
        Exit Sub
    End If

    ' ここまででスポットデータの取り込みは成功している。以降のフォワード処理が
    ' 失敗しても、その日のスポット取り込み自体を失敗扱いにはしない。
    Dim forwardSyncMsg As String
    forwardSyncMsg = ForwardSpec.RunForwardSpecSync_Silent()
    LogStatus "ForwardSpec同期: " & forwardSyncMsg

    Dim forwardRefreshMsg As String
    forwardRefreshMsg = ForwardHistory.RunForwardHistoryRefresh_Silent()
    LogStatus "ForwardHistory更新: " & forwardRefreshMsg

    If InStr(1, forwardSyncMsg, "エラー", vbTextCompare) = 1 Or _
       InStr(1, forwardRefreshMsg, "エラー", vbTextCompare) = 1 Then
        LogStatus "完了(フォワード処理に警告あり): DB取り込みは成功しました"
        NotifyIfInteractive silent, "更新は完了しましたが、フォワードスワップレートの処理に" & _
            "問題がありました。UpdateLogシートを確認してください。", vbExclamation
    Else
        LogStatus "完了: DB取り込み・フォワード更新ともに成功しました"
        NotifyIfInteractive silent, "更新が完了しました。", vbInformation
    End If
    Exit Sub

ConfigError:
    LogStatus "中止: データシート名の取得に失敗しました(" & Err.Description & ")"
    NotifyIfInteractive silent, "設定エラー: " & Err.Description, vbExclamation
    Exit Sub

SaveError:
    LogStatus "中止: ワークブックの保存に失敗しました(" & Err.Description & ")"
    NotifyIfInteractive silent, "ワークブックの保存に失敗しました: " & Err.Description, vbExclamation
End Sub


' Controlシート(Control!B24)からデータシート名を読む。
' シートが無い/値が空の場合はエラーを発生させ、呼び出し元(RunDailyUpdate)で中断する。
Private Function GetDataSheetName() As String
    Dim ctrlWs As Worksheet
    Dim val As String

    On Error Resume Next
    Set ctrlWs = ThisWorkbook.Sheets(CONTROL_SHEET)
    On Error GoTo 0
    If ctrlWs Is Nothing Then
        Err.Raise vbObjectError + 1, , "シート '" & CONTROL_SHEET & "' が見つかりません。"
    End If

    val = Trim(ctrlWs.Range(DATA_SHEET_NAME_CELL).Value)
    If val = "" Then
        Err.Raise vbObjectError + 2, , CONTROL_SHEET & "!" & DATA_SHEET_NAME_CELL & " にデータシート名が入力されていません。"
    End If

    GetDataSheetName = val
End Function


' 完全再計算を実行し、Application.CalculationStateがxlDoneになるまで待つ。
' Bloombergの非同期更新(BDH確定後の内部プッシュ)に備え、完了後さらに2秒待つ。
' timeoutSec以内に終わらなければFalseを返す(無限待機にしない)。
Private Function WaitForRecalculation(timeoutSec As Long) As Boolean
    Dim startTime As Double

    Application.CalculateFullRebuild
    startTime = Timer

    Do While Application.CalculationState <> xlDone
        DoEvents
        If Timer - startTime > timeoutSec Then
            WaitForRecalculation = False
            Exit Function
        End If
    Loop

    Application.Wait Now + TimeValue("0:00:02")
    WaitForRecalculation = True
End Function


' startCellRefから始まるデータ範囲(そのシートの最終行・最終列まで)を走査し、
' 最初に見つかったエラー値セルのアドレスを返す(無ければ空文字)。
Private Function FindErrorCell(sheetName As String, startCellRef As String) As String
    Dim ws As Worksheet
    Dim startCell As Range
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        FindErrorCell = "(シート'" & sheetName & "'が見つかりません)"
        Exit Function
    End If

    Set startCell = ws.Range(startCellRef)
    lastRow = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp).Row
    lastCol = ws.Cells(startCell.Row, ws.Columns.Count).End(xlToLeft).Column

    For r = startCell.Row To lastRow
        For c = startCell.Column To lastCol
            If IsError(ws.Cells(r, c).Value) Then
                FindErrorCell = ws.Cells(r, c).Address(False, False)
                Exit Function
            End If
        Next c
    Next r

    FindErrorCell = ""
End Function


Private Sub NotifyIfInteractive(silent As Boolean, msg As String, buttons As VbMsgBoxStyle)
    If Not silent Then MsgBox msg, buttons
End Sub


' UpdateLogシート(無ければ自動作成)にタイムスタンプ付きで1行追記する。
Private Sub LogStatus(msg As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = LOG_SHEET
        ws.Range("A1").Value = "Timestamp"
        ws.Range("B1").Value = "Message"
        ws.Columns("A").ColumnWidth = 18
        ws.Columns("B").ColumnWidth = 90
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = msg
End Sub
