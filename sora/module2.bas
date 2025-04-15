Sub STRabc()
    Dim selectedFolder As String
    Dim fso As Object, mainFolder As Object, subFolder As Object
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim STRReportsPath As String
    Application.ScreenUpdating = False
    
    ' Set worksheet and starting row
    Set ws = ThisWorkbook.Sheets(2)
    currentRow = 1

    ' Let user select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Main Folder"
        If .Show <> -1 Then Exit Sub
        selectedFolder = .SelectedItems(1)
    End With

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fso.GetFolder(selectedFolder)
    ws.Cells(2, 1).Value = "Type"
    ws.Cells(3, 1).Value = "Month"
    ' Loop through each subfolder
    For Each subFolder In mainFolder.SubFolders
        ' Write subfolder name into column A
        ws.Cells(currentRow + 3, 1).Value = subFolder.Name
        
        ' Path to STR Reports folder
        STRReportsPath = subFolder.Path & "\STR Reports"
        
        ' Check if STR Reports folder exists
        If fso.FolderExists(STRReportsPath) Then
            ' Call your existing STRmerged procedure here
            Call STRmerged(STRReportsPath, currentRow)
        Else
            ws.Cells(currentRow, 2).Value = "STR Reports folder not found"
        End If

        ' Move to the next row block (skip 3 rows)
        currentRow = currentRow + 3
    Next subFolder
    Application.ScreenUpdating = True
    
    With ws
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For i = lastRow To 1 Step -1
            If Trim(.Cells(i, "A").Value) = "" Then
                .Rows(i).Delete
            End If
        Next i
    End With

End Sub

Sub STRmerged(folderPath As String, targetRow As Long)

    Dim fileName As String
    Dim externalWb As Workbook
    Dim compSheet As Worksheet
    Dim mainSheet As Worksheet
    Dim currentWb As Workbook
    Dim destCol As Long
    Dim col As Long
    Dim headerValue As String, cellValue As String
    Dim concatValue As Variant
    Dim occDict As Object, adrDict As Object, revparDict As Object
    Dim i As Long, j As Long
    
    Set currentWb = ThisWorkbook

    On Error Resume Next
    Set mainSheet = currentWb.Sheets("Main")
    On Error GoTo 0

    If mainSheet Is Nothing Then
        MsgBox "Main sheet not found in current workbook. Please create a Main sheet before running this macro."
        Exit Sub
    End If

    fileName = Dir(folderPath & "\*.xls*")
    Set occDict = CreateObject("Scripting.Dictionary")
    Set adrDict = CreateObject("Scripting.Dictionary")
    Set revparDict = CreateObject("Scripting.Dictionary")

    ' Read all data from files
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            Set externalWb = Workbooks.Open(folderPath & "\" & fileName, ReadOnly:=True)

            On Error Resume Next
            Set compSheet = externalWb.Sheets("Comp")
            On Error GoTo 0

            If Not compSheet Is Nothing Then
                ' Process for OCC values (row 21)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    
                    If Not occDict.Exists(concatValue) Then
                        occDict(concatValue) = compSheet.Cells(21, col).Value
                    End If
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    If Not occDict.Exists(concatValue) Then
                        occDict(concatValue) = compSheet.Cells(21, col).Value
                    End If
                Next col
                
                ' Process for ADR values (row 33)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    
                    If Not adrDict.Exists(concatValue) Then
                        adrDict(concatValue) = compSheet.Cells(33, col).Value
                    End If
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    If Not adrDict.Exists(concatValue) Then
                        adrDict(concatValue) = compSheet.Cells(33, col).Value
                    End If
                Next col
                
                ' Process for RevPAR values (row 45)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    
                    If Not revparDict.Exists(concatValue) Then
                        revparDict(concatValue) = compSheet.Cells(45, col).Value
                    End If
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    If Not revparDict.Exists(concatValue) Then
                        revparDict(concatValue) = compSheet.Cells(45, col).Value
                    End If
                Next col
            End If

            externalWb.Close False
        End If
        fileName = Dir
    Loop

    ' Set starting column (B = 2, since A has subfolder name)
    destCol = 2
    
    ' Write OCC
    ProcessSTRocc occDict, mainSheet, destCol, targetRow
    
    ' Leave a column empty
    destCol = destCol + 1
    
    ' Write ADR
    ProcessSTRadr adrDict, mainSheet, destCol, targetRow
    
    ' Leave a column empty
    destCol = destCol + 1
    
    ' Write RevPAR
    ProcessSTRrevpar revparDict, mainSheet, destCol, targetRow

End Sub


Sub STRmergedWorking()

    Dim folderPath As String
    Dim fileName As String
    Dim externalWb As Workbook
    Dim compSheet As Worksheet
    Dim mainSheet As Worksheet
    Dim currentWb As Workbook
    Dim destCol As Long
    Dim col As Long
    Dim headerValue As String, cellValue As String
    Dim concatValue As Variant
    Dim occDict As Object, adrDict As Object, revparDict As Object
    Dim i As Long, j As Long
    
    Set currentWb = ThisWorkbook

    On Error Resume Next
    Set mainSheet = currentWb.Sheets("Main")
    On Error GoTo 0

    If mainSheet Is Nothing Then
        MsgBox "Main sheet not found in current workbook. Please create a Main sheet before running this macro."
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Excel Files"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Operation cancelled."
            Exit Sub
        End If
    End With

    fileName = Dir(folderPath & "\*.xls*")
    Set occDict = CreateObject("Scripting.Dictionary")
    Set adrDict = CreateObject("Scripting.Dictionary")
    Set revparDict = CreateObject("Scripting.Dictionary")

    ' Read all data from files
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            Set externalWb = Workbooks.Open(folderPath & "\" & fileName, ReadOnly:=True)

            On Error Resume Next
            Set compSheet = externalWb.Sheets("Comp")
            On Error GoTo 0

            If Not compSheet Is Nothing Then
                ' Process for OCC values (row 21)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    
                    If Not occDict.Exists(concatValue) Then
                        occDict(concatValue) = compSheet.Cells(21, col).Value
                    End If
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    If Not occDict.Exists(concatValue) Then
                        occDict(concatValue) = compSheet.Cells(21, col).Value
                    End If
                Next col
                
                ' Process for ADR values (row 33)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    
                    If Not adrDict.Exists(concatValue) Then
                        adrDict(concatValue) = compSheet.Cells(33, col).Value
                    End If
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    If Not adrDict.Exists(concatValue) Then
                        adrDict(concatValue) = compSheet.Cells(33, col).Value
                    End If
                Next col
                
                ' Process for RevPAR values (row 45)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    
                    If Not revparDict.Exists(concatValue) Then
                        revparDict(concatValue) = compSheet.Cells(45, col).Value
                    End If
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    If Not revparDict.Exists(concatValue) Then
                        revparDict(concatValue) = compSheet.Cells(45, col).Value
                    End If
                Next col
            End If

            externalWb.Close False
        End If
        fileName = Dir
    Loop

    ' Process OCC values first
    destCol = 1
    
    ' Call original STRocc logic - just the sorting and writing part
    ProcessSTRocc occDict, mainSheet, destCol
    
    ' Leave a column empty
    destCol = destCol + 1
    
    ' Call original STRadr logic - just the sorting and writing part
    ProcessSTRadr adrDict, mainSheet, destCol
    
    ' Leave a column empty
    destCol = destCol + 1
    
    ' Call original STRadr logic - just the sorting and writing part
    ProcessSTRrevpar revparDict, mainSheet, destCol

    MsgBox "All done! OCC and ADR data sorted and written to the Main sheet."

End Sub

' Original STRocc logic for sorting and writing
Private Sub ProcessSTRocc(valueDict As Object, mainSheet As Worksheet, ByRef destCol As Long, targetRow As Long)
    Dim tempList As collection
    Dim yearHeaders As Object
    Dim concatValue As Variant
    Dim valParts() As String, datePart As String
    Dim dateArray() As Variant
    Dim keyArray() As Variant
    Dim yearKeys() As Variant, yearValues() As Variant
    Dim k As Long, i As Long, j As Long
    Dim dTemp As Variant, kTemp As Variant
    Dim tempYear As Variant, tempKey As Variant
    
    Set tempList = New collection
    Set yearHeaders = CreateObject("Scripting.Dictionary")

    ' Separate into tempList and yearHeaders
    For Each concatValue In valueDict.Keys
        valParts = Split(concatValue, "-")
        datePart = Trim(valParts(0))
        
        If IsDate(concatValue) Then
            tempList.Add concatValue
        ElseIf IsNumeric(datePart) And Len(datePart) = 4 Then
            yearHeaders(concatValue) = valueDict(concatValue)
        Else
            tempList.Add concatValue
        End If
    Next

    ' Sort values by date properly
    ReDim dateArray(1 To tempList.Count)
    ReDim keyArray(1 To tempList.Count)

    k = 1
    For Each concatValue In tempList
        If IsDate(Split(concatValue, " ")(0)) Then
            dateArray(k) = CDate(Split(concatValue, " ")(0))
            keyArray(k) = concatValue
            k = k + 1
        End If
    Next

    If k > 1 Then
        ReDim Preserve dateArray(1 To k - 1)
        ReDim Preserve keyArray(1 To k - 1)

        ' Sort using bubble sort
        For i = LBound(dateArray) To UBound(dateArray) - 1
            For j = i + 1 To UBound(dateArray)
                If dateArray(i) > dateArray(j) Then
                    dTemp = dateArray(i)
                    dateArray(i) = dateArray(j)
                    dateArray(j) = dTemp
                    
                    kTemp = keyArray(i)
                    keyArray(i) = keyArray(j)
                    keyArray(j) = kTemp
                End If
            Next j
        Next i

        ' Write sorted date values with occ in row 1, date in row 2, and value in row 3
        For i = 1 To UBound(keyArray)
            concatValue = keyArray(i)
            mainSheet.Cells(1 + targetRow, destCol).Value = "occ" ' Put occ in row 1
            mainSheet.Cells(2 + targetRow, destCol).Value = concatValue ' Put date in row 2
            mainSheet.Cells(3 + targetRow, destCol).Value = valueDict(concatValue) ' Put value in row 3
            destCol = destCol + 1
        Next i
    End If

    ' ----------- Sort & Write Year Headers -------------
    If yearHeaders.Count > 0 Then
        ReDim yearKeys(1 To yearHeaders.Count)
        ReDim yearValues(1 To yearHeaders.Count)

        i = 1
        For Each concatValue In yearHeaders.Keys
            valParts = Split(concatValue, "-")
            yearKeys(i) = CLng(valParts(0))
            yearValues(i) = concatValue
            i = i + 1
        Next

        ' Sort yearKeys and yearValues accordingly
        For i = LBound(yearKeys) To UBound(yearKeys) - 1
            For j = i + 1 To UBound(yearKeys)
                If yearKeys(i) > yearKeys(j) Then
                    tempYear = yearKeys(i)
                    yearKeys(i) = yearKeys(j)
                    yearKeys(j) = tempYear
                    
                    tempKey = yearValues(i)
                    yearValues(i) = yearValues(j)
                    yearValues(j) = tempKey
                End If
            Next j
        Next i

        ' Write year headers with occ in row 1, year info in row 2, value in row 3
        For i = LBound(yearValues) To UBound(yearValues)
            mainSheet.Cells(1 + targetRow, destCol).Value = "occ" ' Put occ in row 1
            mainSheet.Cells(2 + targetRow, destCol).Value = yearValues(i) ' Put year info in row 2
            mainSheet.Cells(3 + targetRow, destCol).Value = yearHeaders(yearValues(i)) ' Put value in row 3
            destCol = destCol + 1
        Next i
    End If
End Sub

' Original STRadr logic for sorting and writing
Private Sub ProcessSTRadr(valueDict As Object, mainSheet As Worksheet, ByRef destCol As Long, targetRow As Long)
    Dim tempList As collection
    Dim yearHeaders As Object
    Dim concatValue As Variant
    Dim valParts() As String, datePart As String
    Dim dateArray() As Variant
    Dim keyArray() As Variant
    Dim yearKeys() As Variant, yearValues() As Variant
    Dim k As Long, i As Long, j As Long
    Dim dTemp As Variant, kTemp As Variant
    Dim tempYear As Variant, tempKey As Variant
    
    Set tempList = New collection
    Set yearHeaders = CreateObject("Scripting.Dictionary")

    ' Separate into tempList and yearHeaders
    For Each concatValue In valueDict.Keys
        valParts = Split(concatValue, "-")
        datePart = Trim(valParts(0))
        
        If IsDate(concatValue) Then
            tempList.Add concatValue
        ElseIf IsNumeric(datePart) And Len(datePart) = 4 Then
            yearHeaders(concatValue) = valueDict(concatValue)
        Else
            tempList.Add concatValue
        End If
    Next

    ' Sort values by date properly
    ReDim dateArray(1 To tempList.Count)
    ReDim keyArray(1 To tempList.Count)

    k = 1
    For Each concatValue In tempList
        If IsDate(Split(concatValue, " ")(0)) Then
            dateArray(k) = CDate(Split(concatValue, " ")(0))
            keyArray(k) = concatValue
            k = k + 1
        End If
    Next

    If k > 1 Then
        ReDim Preserve dateArray(1 To k - 1)
        ReDim Preserve keyArray(1 To k - 1)

        ' Sort using bubble sort
        For i = LBound(dateArray) To UBound(dateArray) - 1
            For j = i + 1 To UBound(dateArray)
                If dateArray(i) > dateArray(j) Then
                    dTemp = dateArray(i)
                    dateArray(i) = dateArray(j)
                    dateArray(j) = dTemp
                    
                    kTemp = keyArray(i)
                    keyArray(i) = keyArray(j)
                    keyArray(j) = kTemp
                End If
            Next j
        Next i

        ' Write sorted date values with adr in row 1, date in row 2, and value in row 3
        For i = 1 To UBound(keyArray)
            concatValue = keyArray(i)
            mainSheet.Cells(1 + targetRow, destCol).Value = "adr"  ' Put adr in row 1
            mainSheet.Cells(2 + targetRow, destCol).Value = concatValue  ' Put date in row 2
            mainSheet.Cells(3 + targetRow, destCol).Value = valueDict(concatValue)  ' Put value in row 3
            destCol = destCol + 1
        Next i
    End If

    ' ----------- Sort & Write Year Headers -------------
    If yearHeaders.Count > 0 Then
        ReDim yearKeys(1 To yearHeaders.Count)
        ReDim yearValues(1 To yearHeaders.Count)

        i = 1
        For Each concatValue In yearHeaders.Keys
            valParts = Split(concatValue, "-")
            yearKeys(i) = CLng(valParts(0))
            yearValues(i) = concatValue
            i = i + 1
        Next

        ' Sort yearKeys and yearValues accordingly
        For i = LBound(yearKeys) To UBound(yearKeys) - 1
            For j = i + 1 To UBound(yearKeys)
                If yearKeys(i) > yearKeys(j) Then
                    tempYear = yearKeys(i)
                    yearKeys(i) = yearKeys(j)
                    yearKeys(j) = tempYear
                    
                    tempKey = yearValues(i)
                    yearValues(i) = yearValues(j)
                    yearValues(j) = tempKey
                End If
            Next j
        Next i

        ' Write year headers with adr in row 1, year info in row 2, value in row 3
        For i = LBound(yearValues) To UBound(yearValues)
            mainSheet.Cells(1 + targetRow, destCol).Value = "adr"  ' Put adr in row 1
            mainSheet.Cells(2 + targetRow, destCol).Value = yearValues(i)  ' Put year info in row 2
            mainSheet.Cells(3 + targetRow, destCol).Value = yearHeaders(yearValues(i))  ' Put value in row 3
            destCol = destCol + 1
        Next i
    End If
End Sub



' Original STRadr logic for sorting and writing
Private Sub ProcessSTRrevpar(valueDict As Object, mainSheet As Worksheet, ByRef destCol As Long, targetRow As Long)
    Dim tempList As collection
    Dim yearHeaders As Object
    Dim concatValue As Variant
    Dim valParts() As String, datePart As String
    Dim dateArray() As Variant
    Dim keyArray() As Variant
    Dim yearKeys() As Variant, yearValues() As Variant
    Dim k As Long, i As Long, j As Long
    Dim dTemp As Variant, kTemp As Variant
    Dim tempYear As Variant, tempKey As Variant
    
    Set tempList = New collection
    Set yearHeaders = CreateObject("Scripting.Dictionary")

    ' Separate into tempList and yearHeaders
    For Each concatValue In valueDict.Keys
        valParts = Split(concatValue, "-")
        datePart = Trim(valParts(0))
        
        If IsDate(concatValue) Then
            tempList.Add concatValue
        ElseIf IsNumeric(datePart) And Len(datePart) = 4 Then
            yearHeaders(concatValue) = valueDict(concatValue)
        Else
            tempList.Add concatValue
        End If
    Next

    ' Sort values by date properly
    ReDim dateArray(1 To tempList.Count)
    ReDim keyArray(1 To tempList.Count)

    k = 1
    For Each concatValue In tempList
        If IsDate(Split(concatValue, " ")(0)) Then
            dateArray(k) = CDate(Split(concatValue, " ")(0))
            keyArray(k) = concatValue
            k = k + 1
        End If
    Next

    If k > 1 Then
        ReDim Preserve dateArray(1 To k - 1)
        ReDim Preserve keyArray(1 To k - 1)

        ' Sort using bubble sort
        For i = LBound(dateArray) To UBound(dateArray) - 1
            For j = i + 1 To UBound(dateArray)
                If dateArray(i) > dateArray(j) Then
                    dTemp = dateArray(i)
                    dateArray(i) = dateArray(j)
                    dateArray(j) = dTemp
                    
                    kTemp = keyArray(i)
                    keyArray(i) = keyArray(j)
                    keyArray(j) = kTemp
                End If
            Next j
        Next i

        ' Write sorted date values with adr in row 1, date in row 2, and value in row 3
        For i = 1 To UBound(keyArray)
            concatValue = keyArray(i)
            mainSheet.Cells(1 + targetRow, destCol).Value = "revpar"  ' Put adr in row 1
            mainSheet.Cells(2 + targetRow, destCol).Value = concatValue  ' Put date in row 2
            mainSheet.Cells(3 + targetRow, destCol).Value = valueDict(concatValue)  ' Put value in row 3
            destCol = destCol + 1
        Next i
    End If

    ' ----------- Sort & Write Year Headers -------------
    If yearHeaders.Count > 0 Then
        ReDim yearKeys(1 To yearHeaders.Count)
        ReDim yearValues(1 To yearHeaders.Count)

        i = 1
        For Each concatValue In yearHeaders.Keys
            valParts = Split(concatValue, "-")
            yearKeys(i) = CLng(valParts(0))
            yearValues(i) = concatValue
            i = i + 1
        Next

        ' Sort yearKeys and yearValues accordingly
        For i = LBound(yearKeys) To UBound(yearKeys) - 1
            For j = i + 1 To UBound(yearKeys)
                If yearKeys(i) > yearKeys(j) Then
                    tempYear = yearKeys(i)
                    yearKeys(i) = yearKeys(j)
                    yearKeys(j) = tempYear
                    
                    tempKey = yearValues(i)
                    yearValues(i) = yearValues(j)
                    yearValues(j) = tempKey
                End If
            Next j
        Next i

        ' Write year headers with adr in row 1, year info in row 2, value in row 3
        For i = LBound(yearValues) To UBound(yearValues)
            mainSheet.Cells(1 + targetRow, destCol).Value = "revpar"  ' Put adr in row 1
            mainSheet.Cells(2 + targetRow, destCol).Value = yearValues(i)  ' Put year info in row 2
            mainSheet.Cells(3 + targetRow, destCol).Value = yearHeaders(yearValues(i)) ' Put value in row 3
            destCol = destCol + 1
        Next i
    End If
End Sub

