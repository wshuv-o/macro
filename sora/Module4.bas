Attribute VB_Name = "Module4"
Sub STRModule4()
    Dim selectedFolder As String
    Dim fso As Object, mainFolder As Object, subFolder As Object
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim STRReportsPath As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet and starting row
    Set ws = ThisWorkbook.Sheets("Main")
    currentRow = 3  ' Starting at row 2
    
    ' Clear existing data
    ws.UsedRange.Clear
    
    ' Set up headers in row 1
    With ws
        .Cells(2, 2).Value = "Folder Name"
        .Cells(2, 3).Value = "UW Property Name"
        .Cells(2, 4).Value = "File Name"
        .Cells(2, 5).Value = "Raw Date"
        .Cells(2, 6).Value = "Date"
        .Cells(2, 7).Value = "Comp 1 Occ"
        .Cells(2, 8).Value = "Date"
        .Cells(2, 9).Value = "Property Occ"
        .Cells(2, 11).Value = "Comp 1 ADR"
        .Cells(2, 12).Value = "Date"
        .Cells(2, 13).Value = "Property ADR"
        .Cells(2, 15).Value = "RevPAR"
        .Cells(2, 16).Value = "Date"
        .Cells(2, 17).Value = "Property RevPAR"
        .Cells(2, 19).Value = "Comp 1 Occ"
        .Cells(2, 20).Value = "Date"
        .Cells(2, 21).Value = "Competitive Set Occ"
        .Cells(2, 23).Value = "Comp 1 ADR"
        .Cells(2, 24).Value = "Date"
        .Cells(2, 25).Value = "Competitive Set ADR"
        .Cells(2, 27).Value = "RevPAR"
        .Cells(2, 28).Value = "Date"
        .Cells(2, 29).Value = "Competitive Set RevPAR"
        .Cells(2, 31).Value = "Comp 1 Occ"
        .Cells(2, 32).Value = "Date"
        .Cells(2, 33).Value = "Index Occ"
        .Cells(2, 35).Value = "Comp 1 ADR"
        .Cells(2, 36).Value = "Date"
        .Cells(2, 37).Value = "Index ADR"
        .Cells(2, 39).Value = "RevPAR"
        .Cells(2, 40).Value = "Date"
        .Cells(2, 41).Value = "Index RevPAR"
        .Cells(2, 43).Value = "Comp 1 Occ"
        .Cells(2, 44).Value = "Date"
        .Cells(2, 45).Value = "Rank Occ"
        .Cells(2, 47).Value = "Comp 1 ADR"
        .Cells(2, 48).Value = "Date"
        .Cells(2, 49).Value = "Rank ADR"
        .Cells(2, 51).Value = "RevPAR"
        .Cells(2, 52).Value = "Date"
        .Cells(2, 53).Value = "Rank RevPAR"
    End With

    ' Let user select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Main Folder"
        If .Show <> -1 Then Exit Sub
        selectedFolder = .SelectedItems(1)
    End With

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mainFolder = fso.GetFolder(selectedFolder)
    
    ' Loop through each subfolder
    For Each subFolder In mainFolder.SubFolders
        ' Path to STR Reports folder
        STRReportsPath = subFolder.Path & "\STR Reports"
        
        ' Check if STR Reports folder exists
        If fso.FolderExists(STRReportsPath) Then
            ' Process the STR Reports folder
            Call ProcessSTRFolder(STRReportsPath, subFolder.Name, ws, currentRow)
        End If
    Next subFolder
    
    ' Format the sheet
    ws.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Processing complete!", vbInformation
End Sub

Sub ProcessSTRFolder(folderPath As String, folderName As String, destSheet As Worksheet, ByRef currentRow As Long)
    Dim fileName As String
    Dim externalWb As Workbook
    Dim compSheet As Worksheet
    Dim dateDict As Object
    Dim dateKeys As Collection
    
    ' Get first file in the folder
    fileName = Dir(folderPath & "\*.xls*")
    
    ' Process each file in the folder
    Do While fileName <> ""
        ' Skip temp files
        If Left(fileName, 2) <> "~$" Then
            ' Open the workbook
            Set externalWb = Workbooks.Open(folderPath & "\" & fileName, ReadOnly:=True)
            Set compSheet = Nothing
            
            ' Find the Comp sheet
            For Each ws In externalWb.Sheets
                If ws.Name Like "Comp*" Then
                    Set compSheet = ws
                    Exit For
                End If
            Next ws
            
            ' If Comp sheet is found, process it
            If Not compSheet Is Nothing Then
                ' Get property name
                Dim propertyName As String
                propertyName = compSheet.Cells(4, 2).Value
                
                ' Collect all unique dates from row 20
                Set dateDict = CreateObject("Scripting.Dictionary")
                Set dateKeys = New Collection
                
                ' Get all date values from row 20, columns C to T and AD to AF
                CollectDates compSheet, dateDict, dateKeys
                
                ' Process each date and write to destination
                Dim dateKey As Variant
                For Each dateKey In dateKeys
                    WriteDataRow compSheet, destSheet, currentRow, dateKey, folderName, fileName, propertyName
                    currentRow = currentRow + 1
                Next dateKey
            End If
            
            ' Close the workbook without saving
            externalWb.Close False
        End If
        
        ' Get next file
        fileName = Dir
    Loop
End Sub

Sub CollectDates(sourceSheet As Worksheet, ByRef dateDict As Object, ByRef dateKeys As Collection)
    Dim col As Long
    Dim headerValue As Variant, cellValue As Variant
    Dim dateKey As Variant
    
    ' Collect dates from columns C to T
    For col = 3 To 20 ' C to T
        ' Get header value (row 19)
        On Error Resume Next
        If sourceSheet.Cells(19, col).MergeArea.Count > 1 Then
            headerValue = sourceSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
        Else
            headerValue = sourceSheet.Cells(19, col).Value
        End If
        On Error GoTo 0
        
        ' Get date value (row 20)
        cellValue = sourceSheet.Cells(20, col).Value
        
        ' Create concatenated key
        If Not IsEmpty(cellValue) Then
            If Not IsEmpty(headerValue) Then
                dateKey = cellValue & "-" & headerValue
            Else
                dateKey = cellValue
            End If
            
            ' Add to collection if not already there
            If Not dateDict.Exists(dateKey) Then
                dateDict.Add dateKey, col
                dateKeys.Add dateKey
            End If
        End If
    Next col
    
    ' Collect dates from columns AD to AF
    For col = 30 To 32 ' AD to AF
        cellValue = sourceSheet.Cells(20, col).Value
        
        If Not IsEmpty(cellValue) Then
            dateKey = cellValue
            
            ' Add to collection if not already there
            If Not dateDict.Exists(dateKey) Then
                dateDict.Add dateKey, col
                dateKeys.Add dateKey
            End If
        End If
    Next col
End Sub

Sub WriteDataRow(sourceSheet As Worksheet, destSheet As Worksheet, ByVal currentRow As Long, ByVal dateKey As Variant, ByVal folderName As String, ByVal fileName As String, ByVal propertyName As String)
    Dim col As Long
    Dim headerValue As Variant, cellValue As Variant
    Dim testKey As Variant
    Dim foundMatch As Boolean
    
    ' Common fields
    With destSheet
        .Cells(currentRow, 2).Value = folderName  ' B: Folder name
        .Cells(currentRow, 4).Value = fileName    ' D: File name
        .Cells(currentRow, 5).Value = propertyName ' E: Property name
        .Cells(currentRow, 6).Value = "=DATEVALUE(INDEX(TEXTSPLIT(INDEX(TEXTSPLIT(E" & currentRow & ",""        ""),2), "": ""), 2))"
    End With
    
    ' Process each column in source sheet to find matching date entries
    ' First for columns C to T
    For col = 3 To 20 ' C to T
        ' Get header value
        On Error Resume Next
        If sourceSheet.Cells(19, col).MergeArea.Count > 1 Then
            headerValue = sourceSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
        Else
            headerValue = sourceSheet.Cells(19, col).Value
        End If
        On Error GoTo 0
        
        ' Get date value
        cellValue = sourceSheet.Cells(20, col).Value
        
        ' Create test key to match against
        If Not IsEmpty(cellValue) Then
            If Not IsEmpty(headerValue) Then
                testKey = cellValue & "-" & headerValue
            Else
                testKey = cellValue
            End If
            
            ' If this column matches our date
            If testKey = dateKey Then
                ' Get values - Property is row 21, CompSet is row 22
                Dim propertyOcc As Variant, compSetOcc As Variant, indexOcc As Variant, rankOcc As Variant
                Dim propertyADR As Variant, compSetADR As Variant, indexADR As Variant, rankADR As Variant
                Dim propertyRevPAR As Variant, compSetRevPAR As Variant, indexRevPAR As Variant, rankRevPAR As Variant
            
                
                ' Occupancy (rows 21 and 22)
                propertyOcc = sourceSheet.Cells(21, col).Value
                compSetOcc = sourceSheet.Cells(22, col).Value
                indexOcc = sourceSheet.Cells(23, col).Value
                rankOcc = sourceSheet.Cells(24, col).Value
                
                ' ADR (rows 33 and 34)
                propertyADR = sourceSheet.Cells(33, col).Value
                compSetADR = sourceSheet.Cells(34, col).Value
                indexADR = sourceSheet.Cells(35, col).Value
                rankADR = sourceSheet.Cells(36, col).Value
                
                ' RevPAR (rows 45 and 46)
                propertyRevPAR = sourceSheet.Cells(45, col).Value
                compSetRevPAR = sourceSheet.Cells(46, col).Value
                indexRevPAR = sourceSheet.Cells(47, col).Value
                rankRevPAR = sourceSheet.Cells(48, col).Value
                
                ' Write to destination
                With destSheet
                    ' Occupancy data
                    .Cells(currentRow, 7).Value = "Comp 1 Occ"    ' G: Type
                    .Cells(currentRow, 8).Value = dateKey         ' H: Date
                    .Cells(currentRow, 9).Value = propertyOcc     ' I: Property Occ
                    .Cells(currentRow, 19).Value = "Comp 1 Occ"   ' S: Type
                    .Cells(currentRow, 20).Value = dateKey        ' T: Date
                    .Cells(currentRow, 21).Value = compSetOcc     ' U: Comp Set Occ
                    .Cells(currentRow, 31).Value = "Comp 1 Occ"   ' S: Type
                    .Cells(currentRow, 32).Value = dateKey        ' T: Date
                    .Cells(currentRow, 33).Value = indexOcc       ' U: Index Occ
                    .Cells(currentRow, 43).Value = "Comp 1 Occ"   ' S: Type
                    .Cells(currentRow, 44).Value = dateKey        ' T: Date
                    .Cells(currentRow, 45).Value = rankOcc       ' U: Index Occ
                    
                    ' ADR data
                    .Cells(currentRow, 11).Value = "Comp 1 ADR"   ' K: Type
                    .Cells(currentRow, 12).Value = dateKey        ' L: Date
                    .Cells(currentRow, 13).Value = propertyADR    ' M: Property ADR
                    .Cells(currentRow, 23).Value = "Comp 1 ADR"   ' W: Type
                    .Cells(currentRow, 24).Value = dateKey        ' X: Date
                    .Cells(currentRow, 25).Value = compSetADR     ' Y: Comp Set ADR
                    .Cells(currentRow, 35).Value = "Comp 1 ADR"   ' W: Type
                    .Cells(currentRow, 36).Value = dateKey        ' X: Date
                    .Cells(currentRow, 37).Value = indexADR       ' Y: Index ADR
                    .Cells(currentRow, 47).Value = "Comp 1 ADR"   ' S: Type
                    .Cells(currentRow, 48).Value = dateKey        ' T: Date
                    .Cells(currentRow, 49).Value = rankADR       ' U: Index Occ
                    
                    ' RevPAR data
                    .Cells(currentRow, 15).Value = "RevPAR"       ' O: Type
                    .Cells(currentRow, 16).Value = dateKey        ' P: Date
                    .Cells(currentRow, 17).Value = propertyRevPAR ' Q: Property RevPAR
                    .Cells(currentRow, 27).Value = "RevPAR"       ' AA: Type
                    .Cells(currentRow, 28).Value = dateKey        ' AB: Date
                    .Cells(currentRow, 29).Value = compSetRevPAR  ' AC: Comp Set RevPAR
                    .Cells(currentRow, 39).Value = "RevPAR"       ' AA: Type
                    .Cells(currentRow, 40).Value = dateKey        ' AB: Date
                    .Cells(currentRow, 41).Value = indexRevPAR  ' AC: Index RevPAR
                    .Cells(currentRow, 51).Value = "RevPAR"   ' S: Type
                    .Cells(currentRow, 52).Value = dateKey        ' T: Date
                    .Cells(currentRow, 53).Value = rankRevPAR       ' U: Index Occ
                End With
                
                foundMatch = True
                Exit For
            End If
        End If
    Next col
    
    ' If not found in C to T, check columns AD to AF
    If Not foundMatch Then
        For col = 30 To 32 ' AD to AF
            cellValue = sourceSheet.Cells(20, col).Value
            
            If Not IsEmpty(cellValue) And cellValue = dateKey Then
                ' Get values - Property is row 21, CompSet is row 22
                Dim propOcc As Variant, cSetOcc As Variant, inOcc As Variant, rnkOcc As Variant
                Dim propADR As Variant, cSetADR As Variant, inADR As Variant, rnkADR As Variant
                Dim propRevPAR As Variant, cSetRevPAR As Variant, inRevPAR As Variant, rnkRevPAR As Variant
                
                ' Occupancy (rows 21 and 22)
                propOcc = sourceSheet.Cells(21, col).Value
                cSetOcc = sourceSheet.Cells(22, col).Value
                inOcc = sourceSheet.Cells(23, col).Value
                rnkOcc = sourceSheet.Cells(24, col).Value
                
                ' ADR (rows 33 and 34)
                propADR = sourceSheet.Cells(33, col).Value
                cSetADR = sourceSheet.Cells(34, col).Value
                inADR = sourceSheet.Cells(35, col).Value
                rnkADR = sourceSheet.Cells(36, col).Value
                
                ' RevPAR (rows 45 and 46)
                propRevPAR = sourceSheet.Cells(45, col).Value
                cSetRevPAR = sourceSheet.Cells(46, col).Value
                inRevPAR = sourceSheet.Cells(47, col).Value
                rnkRevPAR = sourceSheet.Cells(48, col).Value
                ' Write to destination
                With destSheet
                    ' Occupancy data
                    .Cells(currentRow, 7).Value = "Comp 1 Occ"   ' G: Type
                    .Cells(currentRow, 8).Value = dateKey        ' H: Date
                    .Cells(currentRow, 9).Value = propOcc        ' I: Property Occ
                    .Cells(currentRow, 19).Value = "Comp 1 Occ"  ' S: Type
                    .Cells(currentRow, 20).Value = dateKey       ' T: Date
                    .Cells(currentRow, 21).Value = cSetOcc       ' U: Comp Set Occ
                    .Cells(currentRow, 31).Value = "Comp 1 Occ"   ' S: Type
                    .Cells(currentRow, 32).Value = dateKey        ' T: Date
                    .Cells(currentRow, 33).Value = inOcc       ' U: Index Occ
                    .Cells(currentRow, 43).Value = "Comp 1 Occ"   ' S: Type
                    .Cells(currentRow, 44).Value = dateKey        ' T: Date
                    .Cells(currentRow, 45).Value = rnkOcc       ' U: Rank Occ
                    
                    ' ADR data
                    .Cells(currentRow, 11).Value = "Comp 1 ADR"  ' K: Type
                    .Cells(currentRow, 12).Value = dateKey       ' L: Date
                    .Cells(currentRow, 13).Value = propADR       ' M: Property ADR
                    .Cells(currentRow, 23).Value = "Comp 1 ADR"  ' W: Type
                    .Cells(currentRow, 24).Value = dateKey       ' X: Date
                    .Cells(currentRow, 25).Value = cSetADR       ' Y: Comp Set ADR
                    .Cells(currentRow, 35).Value = "Comp 1 ADR"   ' W: Type
                    .Cells(currentRow, 36).Value = dateKey        ' X: Date
                    .Cells(currentRow, 37).Value = inADR       ' Y: Index ADR
                    .Cells(currentRow, 47).Value = "Comp 1 ADR"   ' S: Type
                    .Cells(currentRow, 48).Value = dateKey        ' T: Date
                    .Cells(currentRow, 49).Value = rnkADR       ' U: Rank ADR
                    
                    ' RevPAR data
                    .Cells(currentRow, 15).Value = "RevPAR"      ' O: Type
                    .Cells(currentRow, 16).Value = dateKey       ' P: Date
                    .Cells(currentRow, 17).Value = propRevPAR    ' Q: Property RevPAR
                    .Cells(currentRow, 27).Value = "RevPAR"      ' AA: Type
                    .Cells(currentRow, 28).Value = dateKey       ' AB: Date
                    .Cells(currentRow, 29).Value = cSetRevPAR    ' AC: Comp Set RevPAR
                    .Cells(currentRow, 39).Value = "RevPAR"       ' AA: Type
                    .Cells(currentRow, 40).Value = dateKey        ' AB: Date
                    .Cells(currentRow, 41).Value = inRevPAR  ' AC: Index RevPAR
                    .Cells(currentRow, 51).Value = "RevPAR"   ' S: Type
                    .Cells(currentRow, 52).Value = dateKey        ' T: Date
                    .Cells(currentRow, 53).Value = rnkRevPAR       ' U: Rank RevPAR
                End With
                
                Exit For
            End If
        Next col
    End If
End Sub
