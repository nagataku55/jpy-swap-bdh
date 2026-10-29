Attribute VB_Name = "ForwardSpec"
Option Explicit

' ForwardConfigシート(StartTenor | ForwardTenor | Active)の内容を
' FastAPIの/forwardspec/syncへPOSTし、DBのforward_specテーブルを完全に置き換える。
'
' ForwardConfigシートが正データ、DB側はその単純なミラーという位置づけ。
' 新しいフォワードスワップ(例: 2Y2Y)を追加したい時は、ここに行を追加してActiveを
' TRUEにしてから同期すればよい(過去分のバックフィルはForwardHistory.bas側で行う)。
' 行を削除せずActiveをFALSEにした場合、forward_history側の過去データは消えない
' (計算対象から外れるだけ)。

Private Const API_URL As String = "http://127.0.0.1:8000/forwardspec/sync"
Private Const CONFIG_SHEET As String = "ForwardConfig"
Private Const DATA_START_ROW As Long = 2
Private Const INDEX_NAME As String = "TONA"
Private Const CCY As String = "JPY"


' シート上のボタンに割り当てる(結果をMsgBoxで表示)
Public Sub RunForwardSpecSync_Manual()
    Dim msg As String
    msg = DoSync()

    If InStr(1, msg, "エラー", vbTextCompare) = 1 Then
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
End Sub

' DailyUpdate.bas等、MsgBoxを出せない無人実行から呼ぶための静音版。
Public Function RunForwardSpecSync_Silent() As String
    RunForwardSpecSync_Silent = DoSync()
End Function

' ForwardConfigシートを読んでFastAPIへPOSTする共通処理。結果メッセージを返す。
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
        DoSync = "エラー: FastAPIサーバーに接続できませんでした(ForwardSpec同期)"
    Else
        DoSync = "ForwardSpec同期完了 (" & n & "件): " & resp
    End If
End Function


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
