Attribute VB_Name = "Module2"
Sub STRModule2()
    Dim selectedFolder As String
    Dim fso As Object, mainFolder As Object, subFolder As Object
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim STRReportsPath As String
    'Application.ScreenUpdating = False
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
            ' Call the STRmerged procedure
            Call STRmerged(STRReportsPath, currentRow, 21, 33, 45)
        Else
            ws.Cells(currentRow, 2).Value = "STR Reports folder not found"
        End If

        ' Move to the next row block (skip 3 rows)
        currentRow = currentRow + 3
    Next subFolder
    
    'Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub STRmerged(folderPath As String, targetRow As Long, mypropocc As Integer, mypropadr As Integer, myproprev As Integer)
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
    Dim i As Long
    
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
            
            For Each ws In externalWb.Sheets
                If ws.Name Like "Comp*" Then
                    Set compSheet = ws
                    Exit For
                End If
            Next ws
            
            If Not compSheet Is Nothing Then
                ' Process for OCC values (row mypropocc)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    occDict(concatValue) = compSheet.Cells(mypropocc, col).Value
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    occDict(concatValue) = compSheet.Cells(mypropocc, col).Value
                Next col
                
                ' Process for ADR values (row mypropadr)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    adrDict(concatValue) = compSheet.Cells(mypropadr, col).Value
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    adrDict(concatValue) = compSheet.Cells(mypropadr, col).Value
                Next col
                
                ' Process for RevPAR values (row myproprev)
                For col = compSheet.Range("C20").Column To compSheet.Range("T20").Column
                    headerValue = compSheet.Cells(19, col).MergeArea.Cells(1, 1).Value
                    cellValue = compSheet.Cells(20, col).Value
                    concatValue = cellValue & "-" & headerValue
                    revparDict(concatValue) = compSheet.Cells(myproprev, col).Value
                Next col

                For col = compSheet.Range("AD20").Column To compSheet.Range("AF20").Column
                    concatValue = compSheet.Cells(20, col).Value
                    revparDict(concatValue) = compSheet.Cells(myproprev, col).Value
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

' Simplified logic for writing OCC data in order
Private Sub ProcessSTRocc(valueDict As Object, mainSheet As Worksheet, ByRef destCol As Long, targetRow As Long)
    Dim concatValue As Variant
    
    For Each concatValue In valueDict.Keys
        mainSheet.Cells(1 + targetRow, destCol).Value = "Comp 1 Occ"
        mainSheet.Cells(2 + targetRow, destCol).Value = concatValue
        mainSheet.Cells(3 + targetRow, destCol).Value = valueDict(concatValue)
        destCol = destCol + 1
    Next concatValue
End Sub

' Simplified logic for writing ADR data in order
Private Sub ProcessSTRadr(valueDict As Object, mainSheet As Worksheet, ByRef destCol As Long, targetRow As Long)
    Dim concatValue As Variant
    
    For Each concatValue In valueDict.Keys
        mainSheet.Cells(1 + targetRow, destCol).Value = "Comp 1 ADR"
        mainSheet.Cells(2 + targetRow, destCol).Value = concatValue
        mainSheet.Cells(3 + targetRow, destCol).Value = valueDict(concatValue)
        destCol = destCol + 1
    Next concatValue
End Sub

' Simplified logic for writing RevPAR data in order
Private Sub ProcessSTRrevpar(valueDict As Object, mainSheet As Worksheet, ByRef destCol As Long, targetRow As Long)
    Dim concatValue As Variant
    
    For Each concatValue In valueDict.Keys
        mainSheet.Cells(1 + targetRow, destCol).Value = "Comp 1 RevPAR"
        mainSheet.Cells(2 + targetRow, destCol).Value = concatValue
        mainSheet.Cells(3 + targetRow, destCol).Value = valueDict(concatValue)
        destCol = destCol + 1
    Next concatValue
End Sub

