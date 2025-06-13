Attribute VB_Name = "Module1"
Sub CountUWFFiles()
    Dim folderDialog As FileDialog
    Dim folderPath As Variant
    Dim fileCount As Long
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim fileSystem As Object

    ' Delete existing sheet if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("UWF File Count").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create results worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "UWF File Count"
    ws.Cells(1, 1).Value = "Folder Path"
    ws.Cells(1, 2).Value = "File Name"
    rowIndex = 2

    ' Folder picker dialog setup
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.AllowMultiSelect = True

    If folderDialog.Show = -1 Then
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        fileCount = 0

        ' Loop through each selected folder and scan recursively
        For Each folderPath In folderDialog.SelectedItems
            fileCount = fileCount + ScanFolderRecursive(fileSystem.GetFolder(folderPath), ws, rowIndex, fileSystem)
        Next folderPath

        MsgBox "Total UWF_ Excel files found: " & fileCount, vbInformation
    Else
        MsgBox "No folders selected.", vbExclamation
    End If
End Sub

Function ScanFolderRecursive(ByVal folder As Object, ByRef ws As Worksheet, ByRef rowIndex As Long, ByVal fileSystem As Object) As Long
    Dim f As Object
    Dim subFolder As Object
    Dim count As Long
    Dim FileName As String
    Dim extension As String

    count = 0

    ' Check files in current folder
    For Each f In folder.Files
        FileName = f.Name
        extension = LCase(fileSystem.GetExtensionName(FileName))

        ' Check if file starts with "UWF_" and extension is a known Excel extension
        If LCase(Left(FileName, 4)) = "uwf_" And IsExcelFile(extension) Then
            ws.Cells(rowIndex, 1).Value = folder.path
            ws.Cells(rowIndex, 2).Value = FileName
            rowIndex = rowIndex + 1
            count = count + 1
        End If
    Next f

    ' Recursively check subfolders
    For Each subFolder In folder.SubFolders
        count = count + ScanFolderRecursive(subFolder, ws, rowIndex, fileSystem)
    Next subFolder

    ScanFolderRecursive = count
End Function

Function IsExcelFile(ext As String) As Boolean
    Select Case ext
        Case "xlsx", "xlsm", "xlsb", "xls", "xltx", "xltm"
            IsExcelFile = True
        Case Else
            IsExcelFile = False
    End Select
End Function


Sub GetRentRollSheetsinwb()
    Dim folderDialog As FileDialog
    Dim folderPath As Variant
    Dim fileSystem As Object
    Dim folder As Object

    ' Folder picker dialog setup
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.AllowMultiSelect = True

    If folderDialog.Show = -1 Then
        Set fileSystem = CreateObject("Scripting.FileSystemObject")

        ' Loop through each selected folder and scan recursively
        For Each folderPath In folderDialog.SelectedItems
            Set folder = fileSystem.GetFolder(folderPath)
            ' Scan all files including subfolders
            ProcessFolderForRentRollSheets folder, fileSystem
        Next folderPath

        MsgBox "All relevant sheets copied successfully.", vbInformation
    Else
        MsgBox "No folders selected.", vbExclamation
    End If
End Sub

Sub ProcessFolderForRentRollSheets(ByVal folder As Object, ByVal fileSystem As Object)
    Dim file As Object
    Dim subFolder As Object
    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim newSheetName As String
    Dim FileName As String
    Dim extension As String
    Dim trimmedFileName As String
    Dim shortBaseName As String
    Dim copiedSheet As Worksheet
    Dim calcMode As XlCalculation

    ' Save current calculation mode
    calcMode = Application.Calculation

    ' Improve performance by disabling certain application features
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.CutCopyMode = False
    Application.DisplayAlerts = False

    ' Process files in folder
    For Each file In folder.Files
        FileName = file.Name
        extension = LCase(fileSystem.GetExtensionName(FileName))

        If LCase(Left(FileName, 4)) = "uwf_" And IsExcelFile(extension) Then
            ' Open the workbook
            Set wbSource = Workbooks.Open(file.path, ReadOnly:=True)

            trimmedFileName = Replace(FileName, "UWF_", "", , , vbTextCompare)
            trimmedFileName = RemoveFileExtension(trimmedFileName)

            ' Get short base name: first 2 "words" separated by space or underscore from trimmedFileName
            shortBaseName = GetFirstTwoWords(trimmedFileName)

            ' Loop through sheets to find those with "Rent Roll" in the name
            For Each ws In wbSource.Worksheets
                If InStr(1, LCase(ws.Name), "rent roll") > 0 Then
                    ' Create base new sheet name: shortBaseName_SheetName
                    newSheetName = shortBaseName & "_" & ws.Name
                    newSheetName = Trim(newSheetName)

                    ' Copy the sheet to the current workbook
                    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                    Set copiedSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

                    ' Attempt to rename the copied sheet with a unique name, with error handling
                    On Error Resume Next
                    copiedSheet.Name = GetUniqueSheetName(newSheetName)
                    If Err.Number <> 0 Then
                        MsgBox "Error renaming sheet: " & Err.Description, vbExclamation
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            Next ws

            wbSource.Close SaveChanges:=False
        End If
    Next file

    ' Process subfolders recursively
    For Each subFolder In folder.SubFolders
        ProcessFolderForRentRollSheets subFolder, fileSystem
    Next subFolder

    ' Restore application settings
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = calcMode
    Application.EnableEvents = True
End Sub

' Function to get the first two words from a string separated by spaces or underscores
Function GetFirstTwoWords(text As String) As String
    Dim words() As String
    Dim sep As String
    Dim firstTwoWords As String

    ' Prefer splitting by underscore if present, else by space
    If InStr(text, "_") > 0 Then
        sep = "_"
    Else
        sep = " "
    End If

    words = Split(text, sep)

    ' Collect first two words, if fewer words exist take what is there
    If UBound(words) >= 1 Then
        firstTwoWords = words(0) & sep & words(1)
    ElseIf UBound(words) = 0 Then
        firstTwoWords = words(0)
    Else
        firstTwoWords = ""
    End If

    GetFirstTwoWords = firstTwoWords
End Function


Function RemoveFileExtension(FileName As String) As String
    Dim pos As Long
    pos = InStrRev(FileName, ".")
    If pos > 0 Then
        RemoveFileExtension = Left(FileName, pos - 1)
    Else
        RemoveFileExtension = FileName
    End If
End Function

Function GetUniqueSheetName(baseName As String) As String
    Dim newName As String
    Dim suffix As Integer
    Dim nameExists As Boolean

    newName = baseName
    suffix = 1
    nameExists = SheetExists(newName)

    Do While nameExists Or Len(newName) > 31
        ' Trim the baseName if it is too long to allow suffix and stay within 31 chars
        If Len(baseName) > 27 Then
            baseName = Left(baseName, 27)
        End If

        newName = baseName & " (" & suffix & ")"
        suffix = suffix + 1
        nameExists = SheetExists(newName)
    Loop

    GetUniqueSheetName = newName
End Function

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function



