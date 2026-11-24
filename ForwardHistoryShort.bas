Attribute VB_Name = "ForwardHistoryShort"
Option Explicit

' 短期セクター用。Control!B26(最大バックフィル日数。長期側と共通)を読み、
' FastAPIの/forwardhistory_short/refreshをPOSTして、欠損している短期フォワード
' スワップレートをまとめて計算させる(計算はサーバー側のforward_history_short.py)。
' 結果はforward_history_shortテーブルに保存される。
'
' 長期側のForwardHistory.basと同じ作りで、送信先だけが違う。受信タイムアウトも
' 同じく5分(初回バックフィル時はカーブ構築が日数分積み重なるため)。

Private Const API_URL As String = "http://127.0.0.1:8000/forwardhistory_short/refresh"
Private Const CONTROL_SHEET As String = "Control"
Private Const MAX_BACKFILL_DAYS_CELL As String = "B26"
Private Const DEFAULT_MAX_BACKFILL_DAYS As Long = 365

' シート上のボタンに割り当てる(結果をMsgBoxで表示し、UpdateLogシートにも記録する)
Public Sub RunForwardHistoryShortRefresh_Manual()
    Dim msg As String
    msg = DoRefresh()
    DailyUpdate.LogStatus "ForwardHistoryShort更新(手動実行): " & msg

    If InStr(1, msg, "エラー", vbTextCompare) = 1 Then
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
End Sub

' DailyUpdate.bas等、MsgBoxを出せない無人実行から呼ぶための静音版。
Public Function RunForwardHistoryShortRefresh_Silent() As String
    RunForwardHistoryShortRefresh_Silent = DoRefresh()
End Function

Private Function DoRefresh() As String
    Dim maxDays As Long
    Dim body As String
    Dim resp As Variant

    maxDays = GetMaxBackfillDays()

    body = "{""max_backfill_days"":" & CStr(maxDays) & "}"
    resp = HttpPostJson(API_URL, body)

    If IsError(resp) Then
        DoRefresh = "エラー: FastAPIサーバーに接続できませんでした(ForwardHistoryShort更新)"
    Else
        DoRefresh = "ForwardHistoryShort更新完了: " & resp
    End If
End Function

' Control!B26(最大バックフィル日数)を読む。未入力/不正値なら既定(365日)を使う。
Private Function GetMaxBackfillDays() As Long
    Dim ctrlWs As Worksheet
    Dim v As Variant

    On Error Resume Next
    Set ctrlWs = ThisWorkbook.Sheets(CONTROL_SHEET)
    On Error GoTo 0
    If ctrlWs Is Nothing Then
        GetMaxBackfillDays = DEFAULT_MAX_BACKFILL_DAYS
        Exit Function
    End If

    v = ctrlWs.Range(MAX_BACKFILL_DAYS_CELL).Value
    If IsNumeric(v) And v > 0 Then
        GetMaxBackfillDays = CLng(v)
    Else
        GetMaxBackfillDays = DEFAULT_MAX_BACKFILL_DAYS
    End If
End Function


' Returns the response body (String) on success, or CVErr(...) on failure
Private Function HttpPostJson(url As String, jsonBody As String) As Variant
    Dim http As Object

    On Error GoTo ErrHandler
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' 受信タイムアウトのみ長め(5分=300000ms)。初回バックフィルが数十秒〜1分程度
    ' かかることがあるため(コメント参照)。
    http.SetTimeouts 2000, 2000, 5000, 300000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send jsonBody

    If http.Status <> 200 Then
        HttpPostJson = CVErr(xlErrValue)
        Exit Function
    End If

    ' responseTextはContent-Typeにcharsetが無いとUTF-8を正しく解釈できず、警告メッセージの
    ' 日本語が文字化けする。バイト列(responseBody)をUTF-8として明示的に読み直す。
    HttpPostJson = Utf8BytesToString(http.responseBody)
    Exit Function

ErrHandler:
    HttpPostJson = CVErr(xlErrNA)
End Function


Private Function Utf8BytesToString(bytes As Variant) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1   ' adTypeBinary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2   ' adTypeText
    stream.Charset = "utf-8"
    Utf8BytesToString = stream.ReadText
    stream.Close
End Function
