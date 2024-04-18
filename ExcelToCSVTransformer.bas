Sub ExcelToCSVTransformer()
    Dim filePath As String
    Dim csvPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range, row As Range
    Dim findText As String, replaceText As String
    Dim link As Variant
    Dim baseName As String
    Dim saveDir As String
    Dim fileNameWithoutExtension As String

    ' ファイル選択ダイアログを表示
    filePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    If filePath = "False" Then Exit Sub ' ユーザーがキャンセルした場合

    Set wb = Workbooks.Open(filePath)
    
    ' ファイル名から拡張子を除去して基本名を取得
    baseName = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
    saveDir = Application.ActiveWorkbook.Path ' 現在のワークブックの場所を保存場所として使用

    ' 特定パターンの動的置換
    For Each ws In wb.Sheets
        For Each row In ws.UsedRange.Rows
            Dim columnAValue As String
            columnAValue = CStr(row.Cells(1, 1).Value)

            For Each cell In row.Cells
                If InStr(cell.Formula, "x=") > 0 And InStr(cell.Formula, ".xlsx]") > 0 Then
                    Dim replacePattern As String, replaceWith As String
                    replacePattern = "[x=0.xlsx]"
                    replaceWith = "[x=" & columnAValue & ".xlsx]"
                    cell.Formula = Replace(cell.Formula, replacePattern, replaceWith)
                End If
            Next cell
        Next row
    Next ws

    ' 任意のテキスト置換
    findText = InputBox("置換前の言葉を入力してください:", "置換前のテキスト")
    replaceText = InputBox("置換後の言葉を入力してください:", "置換後のテキスト")

    If findText <> "" And replaceText <> "" Then
        For Each ws In wb.Sheets
            For Each cell In ws.UsedRange
                If InStr(cell.Formula, findText) > 0 Then
                    cell.Replace What:=findText, Replacement:=replaceText, LookAt:=xlPart
                End If
            Next cell
        Next ws
    End If

    ' 外部リンクの自動更新
    For Each link In wb.LinkSources(Type:=xlLinkTypeExcelLinks)
        wb.UpdateLink Name:=link, Type:=xlLinkTypeExcelLinks
    Next link

    ' CSV形式で保存（シートは1枚のみと仮定）
    fileNameWithoutExtension = baseName
    csvPath = saveDir & "\" & fileNameWithoutExtension & ".csv"
    wb.Sheets(1).SaveAs Filename:=csvPath, FileFormat:=xlCSV, CreateBackup:=False

    wb.Close SaveChanges:=False

    MsgBox "CSVファイルはこちらに保存されました: " & csvPath, vbInformation, "保存場所"
End Sub

