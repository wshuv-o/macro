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
    With Application.FileDialog(msoFileDialogFolderPicker)
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
                ws.Cells(cellRow, 1).Value = fileName
                ws.Cells(cellRow, 2).Value = rootPath
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
    With Application.FileDialog(msoFileDialogFolderPicker)
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
        fileName = ws.Cells(i, "A").Value
        rootPath = ws.Cells(i, "B").Value ' Get the root path from column B
        
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
