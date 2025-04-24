Attribute VB_Name = "Module3"
Sub STRModule1()
    Dim selectedFolder As String
    Dim fso As Object, mainFolder As Object, subFolder As Object
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim STRReportsPath As String
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet and starting row
    Set ws = ThisWorkbook.Sheets("Main")
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
    
    ' Loop through each subfolder (property)
    For Each subFolder In mainFolder.SubFolders
        ' Write property name in column A
        ws.Cells(currentRow, 1).Value = subFolder.Name
        
        ' Path to STR Reports folder
        STRReportsPath = subFolder.Path & "\STR Reports"
        
        ' Check if STR Reports folder exists
        If fso.FolderExists(STRReportsPath) Then
            ' Call the STRmerged procedure
            Call STRmerged(STRReportsPath, currentRow)
        Else
            ws.Cells(currentRow, 2).Value = "STR Reports folder not found"
            currentRow = currentRow + 1
        End If
    Next subFolder
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationManual
End Sub

Sub STRmerged(folderPath As String, ByRef startRow As Long)
    Dim fileName As String
    Dim externalWb As Workbook
    Dim compSheet As Worksheet
    Dim mainSheet As Worksheet
    Dim currentWb As Workbook
    Dim destRow As Long, destCol As Long
    Dim col As Long
    Dim headerValue As String, cellValue As String
    Dim concatValue As Variant
    Dim occDict As Object, adrDict As Object, revparDict As Object
    
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
    destRow = startRow

    ' Read all data from files
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            ' Write file name in column B for this block of rows
            mainSheet.Cells(destRow, 2).Value = fileName
            mainSheet.Cells(destRow + 1, 2).Value = fileName
            mainSheet.Cells(destRow + 2, 2).Value = fileName
            
            Set externalWb = Workbooks.Open(folderPath & "\" & fileName, ReadOnly:=True)
            
            For Each ws In externalWb.Sheets
                If ws.Name Like "Comp*" Then
                    Set compSheet = ws
                    Exit For
                End If
            Next ws
            
            If Not compSheet Is Nothing Then
                ' Process for OCC values (row 21)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    occDict(concatValue) = compSheet.Cells(21, col).Value
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    occDict(concatValue) = compSheet.Cells(21, col).Value
                Next col
                
                ' Process for ADR values (row 33)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    adrDict(concatValue) = compSheet.Cells(33, col).Value
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    adrDict(concatValue) = compSheet.Cells(33, col).Value
                Next col
                
                ' Process for RevPAR values (row 45)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    revparDict(concatValue) = compSheet.Cells(45, col).Value
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    revparDict(concatValue) = compSheet.Cells(45, col).Value
                Next col
            End If

            externalWb.Close False
            
            ' Write OCC, ADR, RevPAR data for this file starting from column C
            destCol = 3 ' Start writing data from column C
            WriteData occDict, mainSheet, destRow, destCol, "Comp 1 Occ"
            WriteData adrDict, mainSheet, destRow + 1, destCol, "Comp 1 ADR"
            WriteData revparDict, mainSheet, destRow + 2, destCol, "Comp 1 RevPAR"
            
            ' Move to the next block of rows (3 rows per file)
            destRow = destRow + 3
        End If
        fileName = Dir
    Loop
    
    ' Update startRow for the next property
    startRow = destRow
End Sub

Private Sub WriteData(valueDict As Object, mainSheet As Worksheet, writeRow As Long, startCol As Long, headerText As String)
    Dim concatValue As Variant
    Dim currentCol As Long
    currentCol = startCol
    
    ' Write header in the first cell
    mainSheet.Cells(writeRow, currentCol).Value = headerText
    currentCol = currentCol + 1
    
    ' Write each key-value pair in adjacent cells
    For Each concatValue In valueDict.Keys
        mainSheet.Cells(writeRow, currentCol).Value = concatValue
        mainSheet.Cells(writeRow, currentCol + 1).Value = valueDict(concatValue)
        currentCol = currentCol + 2 ' Move to the next column pair
    Next concatValue
End Sub

