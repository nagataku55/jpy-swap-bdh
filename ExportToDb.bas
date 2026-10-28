Attribute VB_Name = "ExportToDb"
Option Explicit

' JPY_TONA_Histシート(ワイド形式: B列=Date, C列以降=テナー別レート, 2行目=テナー
' ラベル, 3行目からデータ)を読み、FastAPIの/historyエンドポイントへJSONでPOSTして
' SQLite(swap_history.db)に取り込む。
'
' レートは%表記のまま送る(サーバー側(history_store.py)が小数に変換して保存する)。
' 事前に `uvicorn main:app` でFastAPIサーバーを起動しておくこと。

Private Const API_URL As String = "http://127.0.0.1:8000/history"
Private Const DATA_SHEET As String = "JPY_TONA_Hist"
Private Const DATE_COL As Long = 2          ' B
Private Const FIRST_TENOR_COL As Long = 3   ' C
Private Const TENOR_LABEL_ROW As Long = 2
Private Const DATA_START_ROW As Long = 3
Private Const INDEX_NAME As String = "TONA"
Private Const CCY As String = "JPY"


Public Sub RunExportToDb()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim parts() As String
    Dim n As Long
    Dim dateStr As String, tenor As String, rateStr As String
    Dim body As String
    Dim resp As Variant

    Set ws = ThisWorkbook.Sheets(DATA_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, DATE_COL).End(xlUp).Row
    lastCol = ws.Cells(TENOR_LABEL_ROW, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < DATA_START_ROW Then
        MsgBox "データ行が見つかりません。", vbExclamation
        Exit Sub
    End If

    ReDim parts(1 To (lastRow - DATA_START_ROW + 1) * (lastCol - FIRST_TENOR_COL + 1))
    n = 0

    For r = DATA_START_ROW To lastRow
        If Not IsEmpty(ws.Cells(r, DATE_COL).Value) Then
            dateStr = Format(ws.Cells(r, DATE_COL).Value, "yyyy-mm-dd")
            For c = FIRST_TENOR_COL To lastCol
                If IsNumeric(ws.Cells(r, c).Value) And Not IsEmpty(ws.Cells(r, c).Value) Then
                    tenor = Trim(ws.Cells(TENOR_LABEL_ROW, c).Value)
                    rateStr = Replace(CStr(ws.Cells(r, c).Value), ",", ".")
                    n = n + 1
                    parts(n) = "{""as_of_date"":""" & dateStr & """,""tenor"":""" & tenor & _
                               """,""rate_pct"":" & rateStr & "}"
                End If
            Next c
        End If
    Next r

    If n = 0 Then
        MsgBox "送信するデータがありません。", vbExclamation
        Exit Sub
    End If

    ReDim Preserve parts(1 To n)

    body = "{""index_name"":""" & INDEX_NAME & """,""ccy"":""" & CCY & _
           """,""source"":""EXCEL_VBA"",""rows"":[" & Join(parts, ",") & "]}"

    resp = HttpPostJson(API_URL, body)

    If IsError(resp) Then
        MsgBox "取り込みに失敗しました。FastAPIサーバー(uvicorn main:app)が起動しているか確認してください。", vbExclamation
    Else
        MsgBox "取り込み完了 (" & n & "件送信): " & resp, vbInformation
    End If
End Sub


' Returns the response body (String) on success, or CVErr(...) on failure
Private Function HttpPostJson(url As String, jsonBody As String) As Variant
    Dim http As Object

    On Error GoTo ErrHandler
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 2000, 2000, 5000, 30000
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
