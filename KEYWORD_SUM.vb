Const INPUT_FILE_PATH As String = "C:" '入力ワークブックのパス
Const INPUT_SHEET_NAME As String = "sheet1" '入力ワークブックのシート名
Const INPUT_FILE_NAME As String = "01.xlsx" '入力ワークブックのファイル名
Const OUTPUT_FILE_NAME As String = "test.xlsm" '出力ワークブックのファイル名
Const OUTPUT_SHEET_NAME As String = "sheet3" '出力ワークブックのシート名

Const TARGET_CULUMN As Long = 2
Dim sourceLastRow As Long
Sub searchTable(wbPath As String)
Set SourceWB = Workbooks.Open(wbPath)
sourceLastRow = Workbooks(INPUT_FILE_NAME).Worksheets(INPUT_SHEET_NAME).Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox sourceLastRow
End Sub
Sub searchKeywordCount(arg1 As Long)
Dim Keyword As String
Dim target_SUM As Long
target_SUM = 0
Keyword = Workbooks(OUTPUT_FILE_NAME).Worksheets(OUTPUT_SHEET_NAME).Range("B3").value
Dim targetRoq_n0 As Variant
Dim targetRow_n1 As Variant
targetRow_n0 = 0
targetRow_n1 = 0
Dim searchingStartRow As Long
searchingStartRow = 1

Dim i As Long
i = 0
Dim targetValue As String
Do While i < sourceLastRow
    On Error GoTo Err_Trap
    targetRow_n1 = WorksheetFunction.Match(Keyword, Range(Cells(searchingStartRow, TARGET_CULUMN), Cells(sourceLastRow, TARGET_CULUMN)), 0)
    searchingStartRow = targetRow_n0 + targetRow_n1 + 1
    targetRow_n0 = searchingStartRow - 1
    targetValue = Cells(searchingStartRow - 1, 3).value
    target_SUM = target_SUM + targetValue
    Debug.Print targetValue
    i = i + 1
Loop

Err_Trap:
    Workbooks(OUTPUT_FILE_NAME).Worksheets(OUTPUT_SHEET_NAME).Range("B4") = target_SUM
End Sub
Sub RUN()
Call searchTable(INPUT_FILE_PATH)
Call searchKeywordCount(sourceLastRow)
End Sub
