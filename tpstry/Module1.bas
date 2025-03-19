Attribute VB_Name = "Module1"
Sub ExtractFilePath()
    Dim folderPath As String
    Dim fileName As String
    Dim rootPath As String
    Dim ws As Worksheet
    Dim cellRow As Integer
    Dim fileSystem As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    
    ' Check if "UW file name" sheet exists, create if not
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UW file name")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "UW file name"
    End If
    On Error GoTo 0
    
    ' Open folder selection dialog
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Source Folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub ' Exit if no folder is selected
        End If
    End With
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Initialize file system objects
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    
    ' Find the next empty row in column A without deleting existing data
    cellRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If cellRow < 2 Then cellRow = 2 ' Ensure at least row 2
    
    For Each subFolder In folder.SubFolders
        ' Loop through each file in the subfolder
        For Each file In subFolder.Files
            fileName = file.Name
            rootPath = subFolder.Path ' Get the root path (parent folder) of the file
            
            If fileName Like "UW*" And _
               (Right(fileName, 4) = ".xls" Or Right(fileName, 5) = ".xlsx" Or Right(fileName, 5) = ".xlsm") Then
                ws.Cells(cellRow, 1).value = fileName
                ws.Cells(cellRow, 2).value = rootPath
                cellRow = cellRow + 1 ' Move to next row
            End If
        Next file
    Next subFolder

    ' Activate the "UW file name" sheet to make sure the user sees the output
    ws.Activate

    Set fileSystem = Nothing
    Set folder = Nothing
    Set subFolder = Nothing
    Set file = Nothing

    MsgBox "File names and root paths extracted successfully!", vbInformation
End Sub

Sub CopyExtractedFiles()
    Dim destFolder As String
    Dim ws As Worksheet
    Dim fileName As String
    Dim rootPath As String
    Dim lastRow As Integer
    Dim i As Integer
    Dim fullFilePath As String
    
    ' Select Destination Folder
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Select Destination Folder"
        If .Show = -1 Then
            destFolder = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Set the sheet containing extracted file names and paths
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UW file name")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'UW file name' not found!", vbExclamation
        Exit Sub
    End If

    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in the "UW file name" sheet and copy the files
    For i = 2 To lastRow ' Assuming file names start from row 2
        fileName = ws.Cells(i, "A").value
        rootPath = ws.Cells(i, "B").value ' Get the root path from column B
        
        ' Ensure the file name and root path are not empty
        If fileName <> "" And rootPath <> "" Then
            fullFilePath = rootPath & "\" & fileName
            If Dir(fullFilePath) <> "" Then
                FileCopy fullFilePath, destFolder & fileName
            Else
                MsgBox "File not found: " & fullFilePath, vbExclamation
            End If
        End If
    Next i

    MsgBox "Files copied successfully!", vbInformation
End Sub

Sub ExtractCopyUWFile()
    ExtractFilePath
    CopyExtractedFiles
End Sub
Sub RemoveRowsWithInvalidStyle()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellStyle As String

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = lastRow To 7 Step -1 ' Loop from bottom to top to avoid skipping rows after deletion
        ' Check if the cell style in column A is not "#_0_E"
        cellStyle = ws.Cells(i, 5).Style ' Assuming you're checking style in column A, you can adjust the column as needed
        If cellStyle <> "#_0_E" Then
            ws.Rows(i).Delete
        End If
    Next i

End Sub



Sub ExtractCashFlowSheets()
    Dim sourceWorkbook As Workbook
    Dim sheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim invalidChars As String
    Dim i As Integer
    Dim sheetCounter As Integer
    Dim folderPath As String
    Dim fileName As String

    ' Select the folder containing .xlsm files
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Folder Containing .xlsm Files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Loop through all the .xlsm files in the selected folder
    fileName = Dir(folderPath & "*.xlsm")
    Application.ScreenUpdating = False

    Do While fileName <> ""
        sheetCounter = 1
        Set sourceWorkbook = Workbooks.Open(folderPath & fileName, ReadOnly:=True)

        For Each sheet In sourceWorkbook.Sheets
            ' Check if the sheet name contains "Cash Flow" but does not contain "Details" or "Footnote"
            If sheet.Name Like "*Cash Flow*" And Not sheet.Name Like "*Aggregate Cash Flow*" And Not sheet.Name Like "*Cash Flow Detail*" And Not sheet.Name Like "*Cash Flow Footnote*" Then
                
                ' Suppress alerts to prevent duplicate named range warnings
                Application.DisplayAlerts = False
                sheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Application.DisplayAlerts = True ' Re-enable alerts

                Set newSheet = ActiveSheet
                newSheetName = newSheet.Range("H5").value

                ' Remove invalid characters from the sheet name
                invalidChars = "/\?*:[]"
                For i = 1 To Len(invalidChars)
                    newSheetName = Replace(newSheetName, Mid(invalidChars, i, 1), "")
                Next i

                ' Trim name length and append counter for uniqueness
                If Len(newSheetName) > 25 Then
                    newSheetName = Left(newSheetName, 25)
                End If
                newSheetName = newSheetName & " (" & sheetCounter & ")"
                sheetCounter = sheetCounter + 1

                ' Attempt renaming with error handling
                On Error Resume Next
                newSheet.Name = newSheetName
                If Err.Number <> 0 Then
                    MsgBox "Error renaming sheet to '" & newSheetName & "'. Please check for invalid characters or length."
                    Err.Clear
                End If
                On Error GoTo 0 ' Reset error handling
            End If
        Next sheet

        sourceWorkbook.Close False
        fileName = Dir
    Loop

    Application.ScreenUpdating = True
    MsgBox "Cash Flow sheets extracted and renamed successfully!", vbInformation
End Sub


Sub ExtractCashFlowSheetsValue()
    Dim sourceWorkbook As Workbook
    Dim sheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim invalidChars As String
    Dim i As Integer
    Dim sheetCounter As Integer
    Dim folderPath As String
    Dim fileName As String
    Dim netCashFlowCell As Range
    Dim copiedRange As Range
    Dim lastRow As Long
    Dim startRow As Long
    Dim dataArray As Variant

    ' Disable pop-up alerts, screen updating, and calculations for faster processing
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Select the folder containing .xlsm files
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Folder Containing .xlsm Files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Loop through all the .xlsm files in the selected folder
    fileName = Dir(folderPath & "*.xlsm")
    
    Do While fileName <> ""
        sheetCounter = 0
        Set sourceWorkbook = Workbooks.Open(folderPath & fileName, ReadOnly:=True)

        For Each sheet In sourceWorkbook.Sheets
            ' Check if the sheet name contains "Cash Flow" but does not contain "Details" or "Footnote"
            If sheet.Name Like "*Cash Flow*" And Not sheet.Name Like "*Aggregate Cash Flow*" And Not sheet.Name Like "*Cash Flow Detail*" And Not sheet.Name Like "*Cash Flow Footnote*" Then

                ' Find the first occurrence of "Net Cash Flow" in column H
                Set netCashFlowCell = Nothing
                For Each cell In sheet.Range("H16:H" & sheet.Cells(sheet.Rows.Count, "H").End(xlUp).Row)
                    If cell.value Like "*Net Cash Flow*" Then
                        Set netCashFlowCell = cell
                        Exit For
                    End If
                Next cell

                ' If we found "Net Cash Flow", set the range dynamically
                If Not netCashFlowCell Is Nothing Then
                    ' Set the range to start at H16 and end at the row where "Net Cash Flow" was found
                    Set copiedRange = sheet.Range("H16:AG" & netCashFlowCell.Row)

                    ' Load data from copiedRange into an array for faster manipulation
                    dataArray = copiedRange.value

                    ' Generate a unique sheet name based on the sheet name
                    newSheetName = sheet.Range("H17").value

                    ' Remove invalid characters from the sheet name
                    invalidChars = "/\?*:[]"
                    For i = 1 To Len(invalidChars)
                        newSheetName = Replace(newSheetName, Mid(invalidChars, i, 1), "")
                    Next i

                    ' Trim name length and append counter for uniqueness
                    If Len(newSheetName) > 25 Then
                        newSheetName = Left(newSheetName, 25)
                    End If

                    ' Check if a sheet with this name already exists, and if so, append a counter
                    Dim tempSheetName As String
                    tempSheetName = newSheetName
                    sheetCounter = 0 ' Reset the counter each time

                    On Error Resume Next
                    Set newSheet = ThisWorkbook.Sheets(tempSheetName)
                    On Error GoTo 0

                     ''If the sheet already exists, append a counter until a unique name is found
                    While Not newSheet Is Nothing
                        sheetCounter = sheetCounter + 1
                        tempSheetName = newSheetName & " (" & sheetCounter & ")"
                        Set newSheet = Nothing
                       On Error Resume Next
                        Set newSheet = ThisWorkbook.Sheets(tempSheetName)
                        On Error GoTo 0
                    Wend

                    ' Now, set the unique sheet name
                    newSheetName = tempSheetName

                    ' Create a new sheet in ThisWorkbook with the unique name
                    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                    newSheet.Name = newSheetName

                    ' Write the dataArray back to the new sheet in one go
                    newSheet.Range("A1").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray

                    ' Copy formats from the original range
                    copiedRange.Copy
                    newSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
                End If
            End If
        Next sheet

        sourceWorkbook.Close False
        fileName = Dir
    Loop

    ' Re-enable pop-up alerts, screen updating, and calculations
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Cash Flow sheets extracted and renamed successfully!", vbInformation
End Sub



