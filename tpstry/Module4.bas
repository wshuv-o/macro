Attribute VB_Name = "Module4"
Sub Action()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellStyle As String
    Dim deleteRows As Range
    
    ' Turn off screen updating, calculations, and events to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        If ws.Range("A1").value = "Cash Flow" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            Set deleteRows = Nothing
            For i = lastRow To 7 Step -1
                cellStyle = ws.Cells(i, 5).Style ' Check style in column E (column 5)
                If cellStyle <> "#_0_E" Then
                    If deleteRows Is Nothing Then
                        Set deleteRows = ws.Rows(i)
                    Else
                        Set deleteRows = Union(deleteRows, ws.Rows(i))
                    End If
                End If
            Next i
        
            ' Delete all rows in one operation
            If Not deleteRows Is Nothing Then
                deleteRows.Delete
            End If
        End If
    Next ws
    
    ' Re-enable screen updating, calculation, and events
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

