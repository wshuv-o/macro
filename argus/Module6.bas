Attribute VB_Name = "Module6"
Sub UpdateRevFormulas()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim cell As Range
    Dim colGG As Long, colGH As Long, colGI As Long, colEM As Long
    Dim r As Long

    colGG = Range("GG1").Column
    colGH = Range("GH1").Column
    colGI = Range("GI1").Column
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
    If ws.Range("A1").value = "Cash Flow" Then

        With ws

            ' Find last used row in the sheet
            On Error Resume Next
            lastRow = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
            On Error GoTo 0
            
            If lastRow = 0 Then GoTo NextSheet ' Skip empty sheets
            
            ' Insert "Rev" into GG9
            .Cells(9, colGG).value = "Rev"
            
            ' Find the last column before GG with data (this will be EM)
            lastCol = colGG - 1
            For col = lastCol To 1 Step -1
                If Application.WorksheetFunction.CountA(.Columns(col)) > 0 Then
                    colEM = col
                    Exit For
                End If
            Next col
            
            ' Fill GG10 to lastRow
            For r = 10 To lastRow
                .Cells(r, colGG).Formula = "=IF(A" & r & "=" & Chr(34) & "Cash Flow Available for Distribution" & Chr(34) & ","""",IF(A" & (r - 1) & "=" & Chr(34) & "Effective Gross Revenue" & Chr(34) & ",""Exp"",GG" & (r - 1) & "))"
            Next r

            ' Fill GH11 to lastRow
            For r = 11 To lastRow
                .Cells(r, colGH).Formula = "=IF(AND(A" & r & "<>"""", " & Cells(r, colEM).Address(False, False) & "=""""),A" & r & ",GH" & (r - 1) & ")"
            Next r
            
            ' Fill GI11 to lastRow
            For r = 11 To lastRow
                .Cells(r, colGI).Formula = "=IF(GG" & r & "="""","""",IF(A" & r & "="""","""",GG" & r & "&""+""&GH" & r & "&""//""&A" & r & "))"
            Next r
            
        End With
        End If
        
NextSheet:
    Next ws


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Formulas inserted successfully in all sheets!"

End Sub


Sub CopyGIFromCashFlowSheetsToCurrent()

    Dim ws As Worksheet
    Dim currentWs As Worksheet
    Dim lastRowSrc As Long, lastRowDest As Long
    Dim colGI As Long
    Dim r As Long

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Set currentWs = ActiveSheet
    colGI = Range("GI1").Column

    ' Start pasting in current sheet column A, from row 1 or after existing data
    lastRowDest = currentWs.Cells(currentWs.Rows.Count, 1).End(xlUp).row
    If lastRowDest < 1 Then lastRowDest = 1 Else lastRowDest = lastRowDest + 1

    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").value = "Cash Flow" Then
            ' Find last row in GI column of this sheet
            On Error Resume Next
            lastRowSrc = ws.Cells(ws.Rows.Count, colGI).End(xlUp).row
            On Error GoTo 0
            
            ' Ensure there is data below row 10 (since data starts from row 11)
            If lastRowSrc >= 11 Then
                ' Copy GI11:GI lastRowSrc values to current sheet column A starting from lastRowDest
                For r = 11 To lastRowSrc
                    currentWs.Cells(lastRowDest, 1).value = ws.Cells(r, colGI).value
                    lastRowDest = lastRowDest + 1
                Next r
            End If
        End If
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Copied GI column from all 'Cash Flow' sheets into current sheet column A."

End Sub

