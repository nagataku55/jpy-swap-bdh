Attribute VB_Name = "ExportHistLevels"
Option Explicit

' config_csvシートの表に列挙されたシートを、それぞれCSVに出力する(データ監査用)。
' 表の開始位置はconfig_csv!G1(例: "A2")のセル参照で指定する(G1自体は固定セル)。
' その開始セルを1列目として、1列目=SheetName, 2列目=SaveFolder,
' 3列目=FileName, 4列目=DataStartCell という並びで読む(行・列とも可変)。
'
' SaveFolderは絶対パス("C:\..."や"\\server\...")でなければ、このブック(ThisWorkbook)が
' 置かれているフォルダからの相対パスとして解釈する。フォルダが存在しなければ自動作成する。
' これにより、PCごとにconfig_csvシートの絶対パスを書き換える必要がなくなる(例: "csv_audit")。
'
' FileNameに "{date}" を含めると、実行日(yyyy-mm-dd)に置き換える。
' これにより、日次の監査ログを上書きせず別ファイルとして残せる。
'
' DataStartCell(例: "B1")のセルがCSVの1行目・1列目(A1)になる
' (それより上の行・左の列は出力に含めない)。未入力の場合はA1からとする。

Private Const CONFIG_SHEET As String = "config_csv"
Private Const JOB_LIST_START_CELL_REF As String = "G1"

Public Sub ExportAllSheetsToCsv()
    Dim jobCount As Long
    Dim resultLines As String
    Dim errMsg As String

    errMsg = RunExportJobs(jobCount, resultLines)
    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation
    ElseIf jobCount = 0 Then
        MsgBox CONFIG_SHEET & " に出力対象の行がありません(" & CONFIG_SHEET & "!" & JOB_LIST_START_CELL_REF & "が指す位置以降にシート名を入力してください)。", vbExclamation
    Else
        MsgBox jobCount & " 件処理しました:" & vbCrLf & resultLines, vbInformation
    End If
End Sub

' DailyUpdate.bas等、MsgBoxを出せない無人実行から呼ぶための静音版。
' 戻り値は結果メッセージ(1行)。呼び出し側でログに書き込む想定。
Public Function ExportAllSheetsToCsv_Quiet() As String
    Dim jobCount As Long
    Dim resultLines As String
    Dim errMsg As String

    errMsg = RunExportJobs(jobCount, resultLines)
    If errMsg <> "" Then
        ExportAllSheetsToCsv_Quiet = "エラー: " & errMsg
    ElseIf jobCount = 0 Then
        ExportAllSheetsToCsv_Quiet = "出力対象の行なし"
    Else
        ExportAllSheetsToCsv_Quiet = jobCount & "件処理: " & Replace(resultLines, vbCrLf, " / ")
    End If
End Function

' config_csvシートのジョブ一覧を読んで各シートをCSV出力する共通処理。
' jobCount/resultLinesをByRefで返し、設定不備があればエラーメッセージを返す(正常時は空文字)。
Private Function RunExportJobs(ByRef jobCount As Long, ByRef resultLines As String) As String
    Dim cfgWs As Worksheet
    Dim jobListStartCell As String
    Dim jobStartRow As Long
    Dim jobStartCol As Long
    Dim r As Long
    Dim sheetName As String
    Dim folderPath As String
    Dim fileName As String
    Dim dataStartCell As String

    jobCount = 0
    resultLines = ""

    On Error Resume Next
    Set cfgWs = ThisWorkbook.Sheets(CONFIG_SHEET)
    On Error GoTo 0
    If cfgWs Is Nothing Then
        RunExportJobs = "シート '" & CONFIG_SHEET & "' が見つかりません。"
        Exit Function
    End If

    jobListStartCell = Trim(cfgWs.Range(JOB_LIST_START_CELL_REF).Value)
    If jobListStartCell = "" Then jobListStartCell = "A2"

    On Error Resume Next
    jobStartRow = cfgWs.Range(jobListStartCell).Row
    jobStartCol = cfgWs.Range(jobListStartCell).Column
    On Error GoTo 0
    If jobStartRow = 0 Then
        RunExportJobs = CONFIG_SHEET & "!" & JOB_LIST_START_CELL_REF & " のセル参照が不正です (" & jobListStartCell & ")"
        Exit Function
    End If

    r = jobStartRow
    Do While Trim(cfgWs.Cells(r, jobStartCol).Value) <> ""
        sheetName = Trim(cfgWs.Cells(r, jobStartCol).Value)
        folderPath = Trim(cfgWs.Cells(r, jobStartCol + 1).Value)
        fileName = Trim(cfgWs.Cells(r, jobStartCol + 2).Value)
        dataStartCell = Trim(cfgWs.Cells(r, jobStartCol + 3).Value)
        If dataStartCell = "" Then dataStartCell = "A1"

        jobCount = jobCount + 1
        resultLines = resultLines & sheetName & ": " & _
                      ExportOneSheet(sheetName, folderPath, fileName, dataStartCell) & vbCrLf

        r = r + 1
    Loop
End Function

' folderPathが絶対パスでなければThisWorkbook.Pathからの相対パスとして解決し、
' 存在しなければ作成する。
Private Function ResolveFolder(folderPath As String) As String
    Dim resolved As String

    If IsAbsolutePath(folderPath) Then
        resolved = folderPath
    Else
        resolved = ThisWorkbook.Path & "\" & folderPath
    End If

    If Right(resolved, 1) <> "\" Then resolved = resolved & "\"

    If Dir(resolved, vbDirectory) = "" Then
        MkDir Left(resolved, Len(resolved) - 1)
    End If

    ResolveFolder = resolved
End Function

Private Function IsAbsolutePath(p As String) As Boolean
    If Len(p) >= 2 Then
        If Mid(p, 2, 1) = ":" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If
    If Left(p, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If
    IsAbsolutePath = False
End Function

' FileName中の"{date}"を実行日(yyyy-mm-dd)に置き換える
Private Function ResolveFileName(fileName As String) As String
    ResolveFileName = Replace(fileName, "{date}", Format(Date, "yyyy-mm-dd"))
End Function

' 1シート分をCSVに出力し、結果メッセージ(成功時はファイルパス、失敗時は理由)を返す
Private Function ExportOneSheet(sheetName As String, folderPath As String, fileName As String, dataStartCell As String) As String
    Dim srcWs As Worksheet
    Dim startCell As Range
    Dim startRow As Long
    Dim startCol As Long
    Dim resolvedFolder As String
    Dim resolvedFileName As String
    Dim fullPath As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim fileNum As Integer
    Dim line As String

    If sheetName = "" Then
        ExportOneSheet = "スキップ(シート名が空です)"
        Exit Function
    End If

    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If srcWs Is Nothing Then
        ExportOneSheet = "エラー: シートが見つかりません"
        Exit Function
    End If

    On Error Resume Next
    Set startCell = srcWs.Range(dataStartCell)
    On Error GoTo 0
    If startCell Is Nothing Then
        ExportOneSheet = "エラー: Data Start Cellが不正です (" & dataStartCell & ")"
        Exit Function
    End If
    startRow = startCell.Row
    startCol = startCell.Column

    If folderPath = "" Or fileName = "" Then
        ExportOneSheet = "エラー: 保存先フォルダまたはファイル名が未入力です"
        Exit Function
    End If

    On Error GoTo FolderError
    resolvedFolder = ResolveFolder(folderPath)
    On Error GoTo 0

    resolvedFileName = ResolveFileName(fileName)
    If LCase(Right(resolvedFileName, 4)) <> ".csv" Then resolvedFileName = resolvedFileName & ".csv"
    fullPath = resolvedFolder & resolvedFileName

    lastRow = srcWs.Cells(srcWs.Rows.Count, startCol).End(xlUp).Row
    lastCol = srcWs.Cells(startRow, srcWs.Columns.Count).End(xlToLeft).Column

    If lastRow < startRow Or lastCol < startCol Then
        ExportOneSheet = "エラー: 出力対象データがありません(Data Start Cell=" & dataStartCell & ")"
        Exit Function
    End If

    fileNum = FreeFile
    Open fullPath For Output As #fileNum

    For r = startRow To lastRow
        line = ""
        For c = startCol To lastCol
            If c > startCol Then line = line & ","
            line = line & CsvField(srcWs.Cells(r, c))
        Next c
        Print #fileNum, line
    Next r

    Close #fileNum

    ExportOneSheet = (lastRow - startRow + 1) & " 行 -> " & fullPath
    Exit Function

FolderError:
    ExportOneSheet = "エラー: 保存先フォルダの解決/作成に失敗しました (" & folderPath & ": " & Err.Description & ")"
End Function

' 1セルをCSVフィールド用の文字列に変換する
' (日付はyyyy-mm-dd、数値はロケール非依存の"."区切り、文字列はカンマ/引用符を含む場合のみ引用符で囲む)
Private Function CsvField(cell As Range) As String
    Dim v As Variant
    v = cell.Value

    If IsEmpty(v) Then
        CsvField = ""
    ElseIf IsDate(v) Then
        CsvField = Format(v, "yyyy-mm-dd")
    ElseIf IsNumeric(v) Then
        CsvField = Replace(CStr(v), ",", ".")
    Else
        Dim s As String
        s = CStr(v)
        If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Then
            CsvField = """" & Replace(s, """", """""") & """"
        Else
            CsvField = s
        End If
    End If
End Function
