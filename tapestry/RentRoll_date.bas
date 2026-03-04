Sub PullRR_AsOfDate()
    '-----------------------------------------------
    Const SEARCH_FOLDER As String = "E:\DD\Loan Review - 2024 - Batch 3\"
    '-----------------------------------------------
    ' Setup
    '-----------------------------------------------
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wsList   As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim fileName As String
    Dim filePath As String
    Dim wbSource As Workbook
    Dim rngValue As Variant

    Set wsList = ThisWorkbook.ActiveSheet
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row

    '-----------------------------------------------
    ' Loop through each file name in Column A
    '-----------------------------------------------
    For i = 2 To lastRow

        fileName = Trim(wsList.Cells(i, 1).Value)
        If fileName = "" Then GoTo NextRow

        ' Search for the file recursively
        filePath = FindFile(SEARCH_FOLDER, fileName)

        If filePath = "" Then
            wsList.Cells(i, 2).Value = "File Not Found"
        Else
            On Error Resume Next
            Set wbSource = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
            On Error GoTo 0

            If wbSource Is Nothing Then
                wsList.Cells(i, 2).Value = "Error Opening File"
            Else
                ' Find first sheet whose name starts with "Rent Roll" (case-insensitive)
                Dim ws       As Worksheet
                Dim wsTarget As Worksheet
                Set wsTarget = Nothing

                For Each ws In wbSource.Sheets
                    If LCase(ws.Name) Like "*) rent roll" Or LCase(ws.Name) Like "rent roll" Then
                        Set wsTarget = ws
                        Exit For  ' Take the first match, ignore others
                    End If
                Next ws

                If wsTarget Is Nothing Then
                    rngValue = "Rent Roll Sheet Not Found"
                Else
                    On Error Resume Next
                    rngValue = wsTarget.Range("K12").Value
                    If Err.Number <> 0 Then
                        rngValue = "Cell Not Found"
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If

                wsList.Cells(i, 2).Value = rngValue

                wbSource.Close SaveChanges:=False
                Set wbSource = Nothing
                Set wsTarget = Nothing
            End If
        End If

NextRow:
    Next i

    '-----------------------------------------------
    ' Restore settings & notify user
    '-----------------------------------------------
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Completed!", vbInformation

End Sub


'================================================================
' Helper: Recursively searches a folder and subfolders for a file
' Returns the full file path if found, or "" if not found
'================================================================
Private Function FindFile(folderPath As String, fileName As String) As String

    Dim fso       As Object
    Dim folder    As Object
    Dim subFolder As Object
    Dim file      As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        FindFile = ""
        Exit Function
    End If

    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If LCase(file.Name) = LCase(fileName) Then
            FindFile = file.Path
            Exit Function
        End If
    Next file

    For Each subFolder In folder.SubFolders
        Dim result As String
        result = FindFile(subFolder.Path, fileName)
        If result <> "" Then
            FindFile = result
            Exit Function
        End If
    Next subFolder

    FindFile = ""

End Function
