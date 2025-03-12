Sub Action()
    Dim ws As Worksheet
    
    ' Turn off screen updating, calculations, and events to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Check if cell A1 contains a value (not empty)
        If ws.Range("A1").value <> "" Then
            ' If A1 has a value, execute the actions
            UnmergeHeaderOnSheet ws
            RemoveRowsWithInvalidStyleOnSheet ws
            'RemoveAmountColumnOnSheet ws
        End If
    Next ws
    
    ' Re-enable screen updating, calculation, and events
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub UnmergeHeaderOnSheet(ws As Worksheet)
    ws.Rows("1:5").UnMerge

    ' Use direct assignment rather than cut and paste to avoid excessive screen updates
    ws.Range("E1").value = ws.Range("D2").value
    ws.Range("J1").value = ws.Range("I2").value
    ws.Range("O1").value = ws.Range("N2").value
    ws.Range("T1").value = ws.Range("S2").value
    ws.Range("Y1").value = ws.Range("X2").value
    ws.Range("E4").value = ws.Range("F4").value
    ws.Range("J4").value = ws.Range("K4").value
    ws.Range("O4").value = ws.Range("P4").value
    ws.Range("T4").value = ws.Range("U4").value
    ws.Range("Y4").value = ws.Range("Z4").value

    ' Optionally, clear the original cells after moving their content (if needed)
    ws.Range("D2, I2, N2, S2, X2, F4, K4, P4, U4, Z4").ClearContents
End Sub

Sub RemoveRowsWithInvalidStyleOnSheet(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellStyle As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through each row starting from row 7 (bottom to top)
    For i = lastRow To 7 Step -1
        cellStyle = ws.Cells(i, 5).Style ' Check style in column E (column 5)
        If cellStyle <> "#_0_E" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub RemoveAmountColumnOnSheet(ws As Worksheet)
    Dim lastCol As Long
    Dim col As Long
    Dim header As String

    ' Find the last used column in the first row (headers)
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through each column starting from column B (right to left to avoid skipping columns after deletion)
    For col = lastCol To 2 Step -1
        header = ws.Cells(5, col).value

        ' Check if the header contains the word "amount" (case-insensitive)
        If Not InStr(1, header, "Amount", vbTextCompare) > 0 Then
            ws.Columns(col).Delete
        End If
    Next col
End Sub

