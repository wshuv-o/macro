Attribute VB_Name = "Module6"
Sub ProcessRentRollAnalysis()
    Dim ws As Worksheet
    Dim trackerSheet As Worksheet
    Dim mfRentRollsSheet As Worksheet
    Dim RentRollSheet As Worksheet
    Dim cellValue As String
    Dim splitValue() As String
    Dim leftPart As String
    Dim rightPart As String
    Dim lastRow As Long
    Dim i As Long
    Dim trackerMatchRow As Variant
    Dim propertyType As String
    Dim loanId As String
    Dim address As String
    Dim mfRentRollsLastRow As Long
    Dim RentRollsLastRow As Long
    Dim mappingSheet As Worksheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    
    Set trackerSheet = ThisWorkbook.Sheets("Tracker")
    Set mfRentRollsSheet = ThisWorkbook.Sheets("MF Rent Rolls")
    Set RentRollSheet = ThisWorkbook.Sheets("Rent Roll")

    If ThisWorkbook.Sheets("Main").Range("Y28").value = "Unmatched" Then
        MsgBox "Please map all types in the Mapping sheet before proceeding."
        ThisWorkbook.Sheets("Mapping").Activate
        Exit Sub
    End If
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Range("A1").value = "Rent Roll Analysis" Then
            cellValue = ws.Range("A2").value
            
            If InStr(cellValue, " ") > 0 Then
                splitValue = Split(cellValue, " ", 2)
                leftPart = splitValue(0)
                rightPart = splitValue(1)
                
                trackerMatchRow = Application.Match(rightPart, trackerSheet.Range("D:D"), 0)
                
                If Not IsError(trackerMatchRow) Then
                    propertyType = trackerSheet.Cells(trackerMatchRow, 9).value
                    loanId = trackerSheet.Cells(trackerMatchRow, 2).value
                    address = trackerSheet.Cells(trackerMatchRow, 5).value
                    
                    'For Multifamily "MF Rent Rolls"
                    If propertyType = "Multifamily" Then
                        mfRentRollsLastRow = mfRentRollsSheet.Cells(mfRentRollsSheet.Rows.Count, 1).End(xlUp).Row
                        If mfRentRollsLastRow < 2 Then mfRentRollsLastRow = 2
                        
                        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                        
                        For i = 15 To lastRow
                            mfRentRollsSheet.Range("A" & mfRentRollsLastRow).value = rightPart
                            mfRentRollsSheet.Range("B" & mfRentRollsLastRow).value = address
                            mfRentRollsSheet.Range("C" & mfRentRollsLastRow).value = loanId
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 4).value = ws.Cells(i, 2).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 5).value = ws.Cells(i, 3).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 7).value = ws.Cells(i, 4).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 8).value = ws.Cells(i, 5).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 9).value = ws.Cells(i, 6).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 10).value = ws.Cells(i, 7).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 11).value = ws.Cells(i, 8).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 12).value = ws.Cells(i, 9).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 13).value = ws.Cells(i, 10).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 14).value = ws.Cells(i, 11).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 15).value = ws.Cells(i, 12).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 16).value = ws.Cells(i, 13).value
                            mfRentRollsSheet.Cells(mfRentRollsLastRow, 17).value = ws.Cells(i, 14).value
                            
                            
                            mfRentRollsLastRow = mfRentRollsLastRow + 1
                            If ws.Cells(i, 1).value = "" Then Exit For
                        Next i
                        
                        
                    'For Commercial "Rent Roll"
                    ElseIf propertyType = "Commercial" Then
                        RentRollsLastRow = RentRollSheet.Cells(RentRollSheet.Rows.Count, 1).End(xlUp).Row
                        If RentRollsLastRow < 3 Then RentRollsLastRow = 3
                        
                        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                        
                        For i = 15 To lastRow
                            RentRollSheet.Range("A" & RentRollsLastRow).value = loanId
                            RentRollSheet.Range("B" & RentRollsLastRow).value = rightPart
                            RentRollSheet.Cells(RentRollsLastRow, 3).value = ws.Cells(i, 2).value
                            RentRollSheet.Cells(RentRollsLastRow, 4).value = ws.Cells(i, 3).value
                            RentRollSheet.Cells(RentRollsLastRow, 5).value = ws.Cells(i, 4).value
                            RentRollSheet.Cells(RentRollsLastRow, 6).value = ws.Cells(i, 5).value
                            RentRollSheet.Cells(RentRollsLastRow, 7).value = ws.Cells(i, 7).value
                            RentRollSheet.Cells(RentRollsLastRow, 8).value = ws.Cells(i, 8).value
                            RentRollSheet.Cells(RentRollsLastRow, 9).value = ws.Cells(i, 25).value
                            RentRollSheet.Cells(RentRollsLastRow, 10).value = ws.Cells(i, 26).value
                            RentRollSheet.Cells(RentRollsLastRow, 11).value = ws.Cells(i, 33).value
                            RentRollSheet.Cells(RentRollsLastRow, 12).value = "=F" & RentRollsLastRow & "*N" & RentRollsLastRow
                            RentRollSheet.Cells(RentRollsLastRow, 13).value = ws.Cells(10, 11).value
                            RentRollSheet.Cells(RentRollsLastRow, 14).value = ws.Cells(i, 11).value
                            
                            
                            RentRollsLastRow = RentRollsLastRow + 1
                            If ws.Cells(i, 1).value = "" Then Exit For
                        Next i
                    End If
                    
                    
                Else
                    MsgBox "Right part not found in Tracker sheet column D. The function expects Property Name in Rent Roll sheet!A2 like this <(num) property_name>"
                End If
            Else
                MsgBox "Cell A2 does not contain a space to split."
            End If
        End If
    Next ws
    
    ' Remove the last row from both Rent Roll and MF Rent Roll sheets after the data insertion
    If mfRentRollsSheet.Cells(mfRentRollsSheet.Rows.Count, 1).End(xlUp).Row > 1 Then
        mfRentRollsSheet.Rows(mfRentRollsSheet.Cells(mfRentRollsSheet.Rows.Count, 1).End(xlUp).Row).Delete
    End If
    
    If RentRollSheet.Cells(RentRollSheet.Rows.Count, 1).End(xlUp).Row > 2 Then
        RentRollSheet.Rows(RentRollSheet.Cells(RentRollSheet.Rows.Count, 1).End(xlUp).Row).Delete
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    
End Sub

