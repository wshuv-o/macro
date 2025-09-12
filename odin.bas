Sub ListUWFilesInLeafFolders()
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim resultRow As Long
    Dim ws As Worksheet
    Dim fileSystem As Object
    Dim rootFolder As Object

    ' Set up the output sheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear
    ws.Range("A1").Value = "Folder Path"
    ws.Range("B1").Value = "UW File(s)"
    resultRow = 2

    ' Open folder picker dialog
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select the Root Folder"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    ' Set up FileSystemObject
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fileSystem.GetFolder(folderPath)

    ' Start recursive search — pass fileSystem object too
    ProcessFolder rootFolder, ws, resultRow, fileSystem

    MsgBox "Done! Check your sheet.", vbInformation
End Sub

Sub ProcessFolder(fld, ws As Worksheet, ByRef resultRow As Long, fileSystem As Object)
    Dim subFld As Object
    Dim fileItem As Object
    Dim uwFiles As Collection
    Dim fileName As String
    Dim fileExt As String
    Dim colIndex As Long

    ' Check if this folder has subfolders
    If fld.SubFolders.Count = 0 Then
        ' It's a leaf folder
        Set uwFiles = New Collection
        For Each fileItem In fld.Files
            fileName = fileItem.Name
            fileExt = LCase(fileSystem.GetExtensionName(fileName))

            If (fileExt = "xls" Or fileExt = "xlsx" Or fileExt = "xlsm") Then
                If Left(fileName, 2) = "UW" Then
                    uwFiles.Add fileItem.Name
                End If
            End If
        Next fileItem

        If uwFiles.Count > 0 Then
            ' Add root folder path
            ws.Cells(resultRow, 1).Value = fld.Path

            ' Add UW filenames
            colIndex = 2
            Dim uwFile As Variant
            For Each uwFile In uwFiles
                ws.Cells(resultRow, colIndex).Value = uwFile
                colIndex = colIndex + 1
            Next uwFile

            resultRow = resultRow + 1
        End If

    Else
        ' Recurse through subfolders
        For Each subFld In fld.SubFolders
            ProcessFolder subFld, ws, resultRow, fileSystem
        Next subFld
    End If
End Sub
