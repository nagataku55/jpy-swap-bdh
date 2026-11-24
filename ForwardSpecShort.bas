Attribute VB_Name = "ForwardSpecShort"
Option Explicit

' 短期セクター用。config_fwd_shortシート(StartTenor | ForwardTenor | Active)の内容を
' FastAPIの/forwardspec_short/syncへPOSTし、DBのforward_spec_shortテーブルを完全に置き換える。
'
' 長期側のForwardSpec.bas(config_fwdシート→forward_spec)と同じ作りで、送信先と
' シート名だけが違う。短期側は将来カーブの補間方法を長期側と変える可能性があるため、
' 計算結果もforward_history_shortという別テーブルに保存される(ForwardHistoryShort.bas)。
'
' 初めて使う時は CreateConfigSheet_Short を1回実行すると、見出しと例の行が入った
' config_fwd_shortシートが作られる(既にある場合は何もしない)。
' 行を削除せずActiveをFALSEにした場合、forward_history_short側の過去データは消えない。

Private Const API_URL As String = "http://127.0.0.1:8000/forwardspec_short/sync"
Private Const CONFIG_SHEET As String = "config_fwd_short"
Private Const DATA_START_ROW As Long = 2
Private Const INDEX_NAME As String = "TONA"
Private Const CCY As String = "JPY"


' シート上のボタンに割り当てる(結果をMsgBoxで表示し、UpdateLogシートにも記録する)
Public Sub RunForwardSpecShortSync_Manual()
    Dim msg As String
    msg = DoSync()
    DailyUpdate.LogStatus "ForwardSpecShort同期(手動実行): " & msg

    If InStr(1, msg, "エラー", vbTextCompare) = 1 Then
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
End Sub

' DailyUpdate.bas等、MsgBoxを出せない無人実行から呼ぶための静音版。
Public Function RunForwardSpecShortSync_Silent() As String
    RunForwardSpecShortSync_Silent = DoSync()
End Function

' config_fwdシートを読んでFastAPIへPOSTする共通処理。結果メッセージを返す。
Private Function DoSync() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim parts() As String
    Dim n As Long
    Dim startTenor As String, fwdTenor As String, activeJson As String
    Dim body As String
    Dim resp As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        DoSync = "エラー: シート '" & CONFIG_SHEET & "' が見つかりません"
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        DoSync = "エラー: " & CONFIG_SHEET & "に設定行がありません"
        Exit Function
    End If

    ReDim parts(1 To lastRow - DATA_START_ROW + 1)
    n = 0

    For r = DATA_START_ROW To lastRow
        startTenor = Trim(ws.Cells(r, 1).Value)
        If startTenor <> "" Then
            fwdTenor = Trim(ws.Cells(r, 2).Value)

            ' Active列: TRUE/FALSE(論理値)でもテキストでも受け付ける
            If CBool(ws.Cells(r, 3).Value) Then
                activeJson = "true"
            Else
                activeJson = "false"
            End If

            n = n + 1
            parts(n) = "{""start_tenor"":""" & startTenor & """,""forward_tenor"":""" & fwdTenor & _
                       """,""index_name"":""" & INDEX_NAME & """,""ccy"":""" & CCY & _
                       """,""active"":" & activeJson & "}"
        End If
    Next r

    If n = 0 Then
        DoSync = "エラー: " & CONFIG_SHEET & "に有効な行がありません(StartTenorが空です)"
        Exit Function
    End If

    ReDim Preserve parts(1 To n)
    body = "[" & Join(parts, ",") & "]"

    resp = HttpPostJson(API_URL, body)

    If IsError(resp) Then
        DoSync = "エラー: FastAPIサーバーに接続できませんでした(ForwardSpecShort同期)"
    Else
        DoSync = "ForwardSpecShort同期完了 (" & n & "件): " & resp
    End If
End Function


' config_fwd_shortシートを新規作成する(既にあれば何もしない)。最初に1回だけ実行する。
' 例の行はあくまで例なので、実際に追跡したいスペックに書き換えてから同期すること。
Public Sub CreateConfigSheet_Short()
    Dim ws As Worksheet
    Dim examples As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        MsgBox "シート '" & CONFIG_SHEET & "' は既にあります。", vbInformation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = CONFIG_SHEET

    ws.Range("A1").Value = "StartTenor"
    ws.Range("B1").Value = "ForwardTenor"
    ws.Range("C1").Value = "Active"
    ws.Range("A1:C1").Font.Bold = True

    ' 例: 1M1M, 3M3M, 6M3M, 6M6M, 9M3M, 1Y3M, 1Y6M, 18M6M
    examples = Array( _
        Array("1M", "1M"), Array("3M", "3M"), Array("6M", "3M"), Array("6M", "6M"), _
        Array("9M", "3M"), Array("1Y", "3M"), Array("1Y", "6M"), Array("18M", "6M"))

    For i = LBound(examples) To UBound(examples)
        ' "1M"等がExcelに日付/数値として解釈されないよう文字列として書き込む
        ws.Cells(DATA_START_ROW + i, 1).NumberFormat = "@"
        ws.Cells(DATA_START_ROW + i, 2).NumberFormat = "@"
        ws.Cells(DATA_START_ROW + i, 1).Value = examples(i)(0)
        ws.Cells(DATA_START_ROW + i, 2).Value = examples(i)(1)
        ws.Cells(DATA_START_ROW + i, 3).Value = True
    Next i

    i = DATA_START_ROW + UBound(examples) + 3
    ws.Cells(i, 1).Value = "※ 短期セクター用(長期はconfig_fwdシート)。行を追加してActiveをTRUEにし、"
    ws.Cells(i + 1, 1).Value = "  ForwardSpecShort.RunForwardSpecShortSync_Manual → ForwardHistoryShort.RunForwardHistoryShortRefresh_Manual"
    ws.Cells(i + 2, 1).Value = "  の順に実行すると、その行の過去分(Control!最大バックフィル日数まで)が計算される。"
    ws.Cells(i + 3, 1).Value = "※ 計算を止めたいだけなら行は消さずActiveをFALSEにする。"
    ws.Range(ws.Cells(i, 1), ws.Cells(i + 3, 1)).Font.Italic = True

    ws.Columns("A:C").ColumnWidth = 14
    MsgBox "シート '" & CONFIG_SHEET & "' を作成しました。例の行を確認・編集してから同期してください。", vbInformation
End Sub


' Returns the response body (String) on success, or CVErr(...) on failure
Private Function HttpPostJson(url As String, jsonBody As String) As Variant
    Dim http As Object

    On Error GoTo ErrHandler
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 2000, 2000, 5000, 15000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send jsonBody

    If http.Status <> 200 Then
        HttpPostJson = CVErr(xlErrValue)
        Exit Function
    End If

    HttpPostJson = http.responseText
    Exit Function

ErrHandler:
    HttpPostJson = CVErr(xlErrNA)
End Function
