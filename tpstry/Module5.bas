Attribute VB_Name = "Module5"
Sub CopyDataToFinancials()
    Dim wsSource As Worksheet
    Dim wsFinancials As Worksheet
    Dim lastRowSource As Long
    Dim lastRowFinancials As Long
    Dim lastColumn As Long
    Dim col As Long
    Dim nextRow As Long
    Dim i As Long
    Dim copyRange As Range
    Dim ws As Worksheet
    Dim flag As Boolean
    
    Dim lastRow As Long
    Dim tempValue As Double
    
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.CutCopyMode = False

    Set wsFinancials = ThisWorkbook.Sheets("Financials")
    nextRow = wsFinancials.Cells(wsFinancials.Rows.Count, "H").End(xlUp).Row + 1
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the sheets you don't want to process
        If ws.Cells(1, 1).value = "Cash Flow" Then
            Set wsSource = ws
            lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

            If IsError(wsSource.Cells(lastRowSource, 1).value) Then
                lastRowSource = lastRowSource - 1
            End If

            lastColumn = wsSource.Cells(5, wsSource.Columns.Count).End(xlToLeft).Column
            
            ' Loop through columns to check for "Amount" in row 5
            For col = 1 To lastColumn
                If wsSource.Cells(5, col).value = "Amount" Then
                    wsFinancials.Cells(nextRow, 1).Formula = "=OFFSET(Tracker!$B$1,MATCH($K" & nextRow & ",Tracker!$D:$D,0)-1,0)"
                    wsFinancials.Range(wsFinancials.Cells(nextRow, 1), wsFinancials.Cells(nextRow + lastRowSource - 7, 1)).FillDown
                    
                    wsFinancials.Cells(nextRow, 6).Formula = "=INDEX('CF Mapping'!$A$3:$A$235, MATCH(H" & nextRow & ", 'CF Mapping'!$C$3:$C$235, 0))"
                    wsFinancials.Range(wsFinancials.Cells(nextRow, 6), wsFinancials.Cells(nextRow + lastRowSource - 7, 6)).FillDown
                    
                    wsFinancials.Cells(nextRow, 7).Formula = "=INDEX('CF Mapping'!$B$3:$B$235, MATCH(H" & nextRow & ", 'CF Mapping'!$C$3:$C$235, 0))"
                    wsFinancials.Range(wsFinancials.Cells(nextRow, 7), wsFinancials.Cells(nextRow + lastRowSource - 7, 7)).FillDown

                    
                    wsFinancials.Cells(nextRow, 12).Formula = "=INDEX(Tracker!$I:$I, MATCH(A" & nextRow & ", Tracker!$B:$B, 0))"
                    wsFinancials.Range(wsFinancials.Cells(nextRow, 12), wsFinancials.Cells(nextRow + lastRowSource - 7, 12)).FillDown
                    
                    
                    wsFinancials.Cells(nextRow, 2).Formula = "=IF(C" & nextRow & "=1,""Excluded"",DATE(YEAR(C" & nextRow & ")-1,MONTH(C" & nextRow & "),DAY(C" & nextRow & ")+1))" ' Populate Column B
                    wsFinancials.Range(wsFinancials.Cells(nextRow, 2), wsFinancials.Cells(nextRow + lastRowSource - 7, 2)).FillDown
                    
                    wsFinancials.Cells(nextRow, 3).Resize(lastRowSource - 6, 1).value = wsSource.Cells(2, col - 1).value                          ' Populate Column C
                    
                    wsFinancials.Cells(nextRow, 5).Resize(lastRowSource - 6, 1).value = wsSource.Cells(4, col + 1).value                          ' Populate Column E
                    
                    Set copyRange = wsSource.Range("A7:A" & lastRowSource)                                                                        ' Populate Column H
                    wsFinancials.Cells(nextRow, 8).Resize(copyRange.Rows.Count, 1).value = copyRange.value
                    
                    Set copyRange = wsSource.Range(wsSource.Cells(7, col), wsSource.Cells(lastRowSource, col))                                    ' Populate Column I
                    wsFinancials.Cells(nextRow, 9).Resize(copyRange.Rows.Count, 1).value = copyRange.value

                    Dim cellValue As String
                    cellValue = wsSource.Cells(2, 1).value ' Get the value from source cashflow[A2]
                    
                    ' Check if first character is '('
                    If Left(cellValue, 1) = "(" Then                                                                                            ' Populate Column K remove (num) from property name if found
                        Dim spacePos As Long
                        spacePos = InStr(1, cellValue, " ") ' Find the position of the first space
                        If spacePos > 0 Then
                            Dim rightPart As String
                            rightPart = Mid(cellValue, spacePos + 1) ' Get the part after the first space
                            wsFinancials.Cells(nextRow, 11).Resize(lastRowSource - 6, 1).value = rightPart
                        Else
                            wsFinancials.Cells(nextRow, 11).Resize(lastRowSource - 6, 1).value = cellValue
                        End If
                    Else
                        wsFinancials.Cells(nextRow, 11).Resize(lastRowSource - 6, 1).value = cellValue
                    End If
                    

                     
                    wsFinancials.Cells(nextRow, 14).Resize(lastRowSource - 6, 1).value = Mid(wsSource.Cells(2, col).value, InStr(1, wsSource.Cells(2, col).value, " ") + 1)  ' Populate Column N (Statement Type)
        
                    wsFinancials.Cells(nextRow, 4).Formula = "=IF(OR(N" & nextRow & "=""Underwriting"", N" & nextRow & "=""Origination""), ""Underwriting"", ""Actual"")"    ' Populate Column D (Statement Type)
                    wsFinancials.Range(wsFinancials.Cells(nextRow, 4), wsFinancials.Cells(nextRow + lastRowSource - 7, 4)).FillDown
        
        
                    ' Now, Insert "Income" or "Expense" based on the "Management Fee"                                                             ' Populate Column M (Line Item Type)
                    flag = False ' Initialize the flag to track if Management Fee is found
                    
                    ' Loop through rows and insert "Income" or "Expense"
                    For i = nextRow To nextRow + lastRowSource - 7
                        If wsFinancials.Cells(i, 8).value = "Management Fee" Then
                            flag = True ' Set flag to True when "Management Fee" is encountered
                        End If

                        ' If Management Fee is found, insert "Expense", otherwise insert "Income"
                        If flag Then
                            wsFinancials.Cells(i, 13).value = "Expense"
                        Else
                            wsFinancials.Cells(i, 13).value = "Income"
                        End If
                    Next i
        
                    nextRow = nextRow + copyRange.Rows.Count
                End If
            Next col
        End If
    Next ws
    
    
    lastRow = wsFinancials.Cells(wsFinancials.Rows.Count, 1).End(xlUp).Row ' Find last row in column A

    For i = 4 To lastRow
        tempValue = wsFinancials.Cells(i, 9).value ' Store Column I value in a variable
        
        If wsFinancials.Cells(i, 8).value Like "Less:*" Then ' Check if Column H contains "Less:*"
            tempValue = -Abs(tempValue) ' Make value negative
        End If
        
        wsFinancials.Cells(i, 9).value = tempValue ' Assign the modified (or unchanged) value back to Column I
    Next i


    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Data copied successfully to the Financials sheet!", vbInformation
End Sub




