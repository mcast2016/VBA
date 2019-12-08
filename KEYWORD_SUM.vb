Const INPUT_FILE_PATH As String = "C:\" '入力ワークブックのパス
Const INPUT_SHEET_NAME As String = "" '入力ワークブックのシート名
Const OUTPUT_FILE_NAME As String = "" '出力ワークブックのファイル名
Const OUTPUT_SHEET_NAME As String = "" '出力ワークブックのシート名
Const SPLIT_DELIMITER As String = "" '区分けしたい文字
Const TARGET_CULUMN As Long = 2

Dim inputFileName As String '入力ワークブックのファイル名
Dim sourceLastRow As Long
Dim fileName(23) As String
Dim targetValue As String
Dim target_SUM As Long
Dim cnt As Long
Sub searchTable(wbPath As String)
Set SourceWB = Workbooks.Open(wbPath & inputFileName, ReadOnly:=True)
'ActiveWindow.Visible = False
sourceLastRow = Workbooks(inputFileName).Worksheets(INPUT_SHEET_NAME).Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox sourceLastRow
End Sub
Sub searchKeywordCount(arg1 As Long)
Dim Keyword As String
Keyword = Workbooks(OUTPUT_FILE_NAME).Worksheets(OUTPUT_SHEET_NAME).Range("B3").Value
Dim targetRoq_n0 As Variant
Dim targetRow_n1 As Variant
targetRow_n0 = 0
targetRow_n1 = 0
Dim searchingStartRow As Long
searchingStartRow = 1

Dim i As Long
i = 0
Do While i < sourceLastRow
    On Error GoTo Err_Trap
    targetRow_n1 = WorksheetFunction.Match(Keyword, Range(Cells(searchingStartRow, TARGET_CULUMN), Cells(sourceLastRow, TARGET_CULUMN)), 0)
    searchingStartRow = targetRow_n0 + targetRow_n1 + 1
    targetRow_n0 = searchingStartRow - 1
    targetValue = Cells(searchingStartRow - 1, 3).Value
    If Len(targetValue) = 3 Then
        Dim targetValueArray() As String
        targetValueArray() = split(targetValue, SPLIT_DELIMITER)
        Dim targetValueArrayint_0 As Long
        Dim targetValueArrayint_1 As Long
        targetValueArrayint_0 = Val(targetValueArray(0))
        targetValueArrayint_1 = Val(targetValueArray(1))
        targetValue = targetValueArrayint_0 + targetValueArrayint_1
    End If
    target_SUM = target_SUM + targetValue
    i = i + 1
Loop

Err_Trap:
    Workbooks(OUTPUT_FILE_NAME).Worksheets(OUTPUT_SHEET_NAME).Range("B4") = target_SUM
End Sub
Sub FILE_FINED()
 Dim buf As String
 cnt = 0
    buf = Dir(INPUT_FILE_PATH & "*.xlsx")
    Do While buf <> ""
        fileName(cnt) = buf
        cnt = cnt + 1
        buf = Dir()
    Loop
End Sub
Sub RUN()
target_SUM = 0
Call FILE_FINED
Dim i As Long
i = 0
Do While cnt - i > 0
    cnt = cnt - 1
    inputFileName = fileName(cnt)
    Call searchTable(INPUT_FILE_PATH)
    Call searchKeywordCount(sourceLastRow)
Loop
End Sub
