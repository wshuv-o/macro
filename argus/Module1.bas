Attribute VB_Name = "Module1"

Sub CombineSheetsFromAllWorkbooksInOneFolder()
    Dim MyPath As String, FilesInPath As String
    Dim MyFiles() As String, SourceRcount As Long, Fnum As Long
    Dim mybook As Workbook, BaseWks As Workbook, sourceRange As Range, destrange As Range
    Dim rnum As Long, CalcMode As Long, Sh As Worksheet
    Dim SaveDriveDir As String, InitFileName As String, fileSaveName As String
    Dim tabName As String
    Dim filePath As String
    
    tabName = ActiveSheet.Range("G2")
    filePath = ActiveSheet.Range("K2")

    'Fill in the path\folder where the Excel files are
    'MyPath = "F:\ODIN\2024\1. January\01.17.2024 (BPREP Logistics_UW Model)_Citigroup_Frank\Argus Pulls\"
    
    MyPath = filePath

    FilesInPath = Dir(MyPath & "*.xl*")
    If FilesInPath = "" Then
        MsgBox "No files found"
        Exit Sub
    End If

    Fnum = 0
    Do While FilesInPath <> ""
        Fnum = Fnum + 1
        ReDim Preserve MyFiles(1 To Fnum)
        MyFiles(Fnum) = FilesInPath
        FilesInPath = Dir()
    Loop

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set BaseWks = ThisWorkbook

    If Fnum > 0 Then
        For Fnum = Fnum To 1 Step -1
            Set mybook = Nothing
            On Error Resume Next
            Set mybook = Workbooks.Open(MyPath & MyFiles(Fnum))
            On Error GoTo 0

            If Not mybook Is Nothing Then
                On Error Resume Next
                Set Sh = mybook.Worksheets(tabName)
                If Not Sh Is Nothing Then
                    Sh.Copy After:=BaseWks.Sheets(BaseWks.Sheets.Count)
                    BaseWks.Sheets(BaseWks.Sheets.Count).Name = Sh.Name & "_" & mybook.Name
                End If
                mybook.Close savechanges:=False
            End If
        Next Fnum
    End If

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub
Sub Fast_CombineSheetsFromAllWorkbooksInOneFolder()
    Dim MyPath As String, fileName As String
    Dim MyFiles() As String, FileCount As Long
    Dim mybook As Workbook, BaseWks As Workbook, Sh As Worksheet
    Dim CalcMode As XlCalculation
    Dim tabName As String
    Dim filePath As String
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim NewName As String
    Dim i As Long

    ' Disable UI updates and calculations
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    ' Cache values instead of calling ActiveSheet.Range multiple times
    Set ws = ActiveSheet
    tabName = ws.Range("G2").value
    filePath = ws.Range("K2").value
    If Right(filePath, 1) <> "\" Then filePath = filePath & "\"

    ' Collect file names
    fileName = Dir(filePath & "*.xl*")
    Do While fileName <> ""
        FileCount = FileCount + 1
        ReDim Preserve MyFiles(1 To FileCount)
        MyFiles(FileCount) = fileName
        fileName = Dir()
    Loop

    If FileCount = 0 Then
        MsgBox "No Excel files found in the folder.", vbExclamation
        GoTo Cleanup
    End If

    Set BaseWks = ThisWorkbook

    ' Loop through each file
    For i = FileCount To 1 Step -1
        On Error Resume Next
        Set mybook = Workbooks.Open(filePath & MyFiles(i), ReadOnly:=True)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        ' Copy the specific sheet if found
        Set Sh = Nothing
        On Error Resume Next
        Set Sh = mybook.Worksheets(tabName)
        On Error GoTo 0

        If Not Sh Is Nothing Then
            Sh.Copy After:=BaseWks.Sheets(BaseWks.Sheets.Count)
            Set NewSheet = BaseWks.Sheets(BaseWks.Sheets.Count)
            ' Sanitize and truncate sheet name if needed
            NewName = Sh.Name & "_" & Left(MyFiles(i), InStrRev(MyFiles(i), ".") - 1)
            If Len(NewName) > 31 Then NewName = Left(NewName, 31)
            On Error Resume Next
            NewSheet.Name = NewName
            On Error GoTo 0
        End If

        mybook.Close savechanges:=False

NextFile:
        Set mybook = Nothing
    Next i

Cleanup:
    ' Re-enable settings
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

Sub ListSheetNamesStartingFromSecond()

    Dim ws As Worksheet
    Dim i As Long
    Dim outputRow As Long

    outputRow = 3 ' Start pasting at B3

    ' Loop through sheets starting from index 2
    For i = 3 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Sheets("_Workings").Range("B" & outputRow).value = ThisWorkbook.Sheets(i).Name
        outputRow = outputRow + 1
    Next i

End Sub
Sub Fast_ListSheetNamesStartingFromSecond()
    Dim wsOutput As Worksheet
    Dim i As Long, countSheets As Long
    Dim outputArr() As String
    Dim outputRow As Long

    ' Performance boosts
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set wsOutput = ThisWorkbook.Sheets("_Workings")
    countSheets = ThisWorkbook.Sheets.Count

    ' Resize output array for only needed rows
    ReDim outputArr(1 To countSheets - 2, 1 To 1) ' Start from sheet 3 onward

    For i = 3 To countSheets
        outputArr(i - 2, 1) = ThisWorkbook.Sheets(i).Name
    Next i

    ' Bulk write to worksheet starting from B3
    wsOutput.Range("B3").Resize(UBound(outputArr, 1), 1).value = outputArr

    ' Re-enable Excel settings
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub


Sub ProcessRightSideColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("_Workings") ' Change this if needed
    
    With ws
        Dim lastRw As Long
        lastRw = .Cells(.Rows.Count, "I").End(xlUp).row
        .Range("I9:I" & lastRw).Copy
        .Range("F9").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
    Application.CutCopyMode = False
    
    Dim startCol As Long
    startCol = ws.Range("HA1").Column ' Start from HA
    
    Dim colIndex As Long
    colIndex = startCol
    
    Dim currentRow As Long
    Dim lastRow As Long
    Dim prevNumber As Variant
    Dim strList As Collection
    Dim insertRow As Long
    Dim i As Long
    
    ' Loop through columns until an empty header cell is found
    Do While ws.Cells(1, colIndex).value <> ""
        currentRow = 9
        lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).row
        
        Do While currentRow <= lastRow
            If IsNumeric(ws.Cells(currentRow, colIndex).value) Then
                prevNumber = ws.Cells(currentRow, colIndex).value
                currentRow = currentRow + 1
            ElseIf Trim(ws.Cells(currentRow, colIndex).value) <> "" Then
                Set strList = New Collection
                Do While Trim(ws.Cells(currentRow, colIndex).value) <> "" And Not IsNumeric(ws.Cells(currentRow, colIndex).value)
                    strList.Add ws.Cells(currentRow, colIndex).value
                    currentRow = currentRow + 1
                Loop
                
                If IsNumeric(ws.Cells(currentRow, colIndex).value) Then
                    insertRow = currentRow
                    For i = strList.Count To 1 Step -1
                        ws.Cells(insertRow, "F").Insert Shift:=xlDown
                        ws.Cells(insertRow, "F").value = strList(i)
                    Next i
                End If
            Else
                Exit Do ' Exit if empty cell in data
            End If
        Loop
        
        colIndex = colIndex + 1
    Loop
    
    MsgBox "Processing completed for all columns to the right of HA!"
End Sub
Sub Fast_DeleteSheets()
    Dim ws As Worksheet
    Dim keyword As String
    Dim i As Long
    Dim workingsSheet As Worksheet
    Dim foundCell As Range
    Dim deleteList As Collection
    Dim sheetName As String

    ' Step 1: Get keyword from active sheet
    keyword = Trim(ActiveSheet.Range("G10").value)
    
    If keyword = "" Then
        MsgBox "Please enter text in cell G10 to specify which sheets to delete.", vbExclamation
        Exit Sub
    End If
    
    If keyword = "Cash Flow" Then
        DeleteBnF ' Your custom procedure
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Step 2: Collect sheets to delete (faster than deleting in loop)
    Set deleteList = New Collection
    For Each ws In ThisWorkbook.Worksheets
        sheetName = ws.Name
        If InStr(1, sheetName, keyword, vbTextCompare) > 0 Then
            deleteList.Add ws
        End If
    Next ws

    ' Step 3: Delete all matching sheets
    For Each ws In deleteList
        ws.Delete
    Next ws

    ' Step 4: Delete columns in "_Workings" sheet
    On Error Resume Next
    Set workingsSheet = ThisWorkbook.Sheets("_Workings")
    On Error GoTo 0

    If Not workingsSheet Is Nothing Then
        With workingsSheet
            Set foundCell = .Columns("B").Find(What:=keyword, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            If Not foundCell Is Nothing Then
                .Columns("F").Delete
                .Columns("B").Delete
            End If
        End With
    End If

    ' Step 5: Restore Excel settings
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Sheets and data matching '" & keyword & "' deleted successfully!", vbInformation
End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim keyword As String
    Dim i As Long
    Dim workingsSheet As Worksheet
    Dim foundCell As Range
    
    

    ' Get the keyword/text from cell G10 on the active sheet
    keyword = ActiveSheet.Range("G10").value
    If keyword = "Cash Flow" Then
        DeleteBnF
    End If
    If Len(keyword) = 0 Then
        MsgBox "Please enter text in cell G10 to specify which sheets to delete."
        Exit Sub
    End If

    Application.DisplayAlerts = False ' Turn off delete confirmation prompts
    
    ' Loop backwards to avoid issues when deleting sheets during the loop
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        ' Check if sheet name contains the keyword (case-insensitive)
        If InStr(1, ws.Name, keyword, vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next i

    ' Set the "_Workings" sheet
    On Error Resume Next
    Set workingsSheet = ThisWorkbook.Worksheets("_Workings")
    On Error GoTo 0

    ' Check if the "_Workings" sheet exists
    If Not workingsSheet Is Nothing Then
        ' Check if the keyword is found in column B of "_Workings"
        Set foundCell = workingsSheet.Columns("B").Find(What:=keyword, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

        ' If found, delete columns B and F
        If Not foundCell Is Nothing Then
            workingsSheet.Columns("B").Delete
            workingsSheet.Columns("F").Delete
        End If
    End If

    Application.DisplayAlerts = True ' Turn prompts back on
End Sub

Sub DeleteBnF()
    Dim ws As Worksheet
    Dim keyword As String
    Dim i As Long
    Dim workingsSheet As Worksheet
    Dim foundCell As Range
    Dim lastRow As Long

    Application.DisplayAlerts = False
    Set workingsSheet = ThisWorkbook.Worksheets("_Workings")
    If Not workingsSheet Is Nothing Then
        ' Check if the keyword is found in column B of "_Workings"
        Set foundCell = workingsSheet.Columns("B").Find(What:=keyword, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

        ' If found, clear the values in columns B and F (not delete columns)
        If Not foundCell Is Nothing Then
            ' Determine last used row in columns B and F
            lastRow = workingsSheet.Cells(workingsSheet.Rows.Count, "B").End(xlUp).row
            workingsSheet.Range("B1:B" & lastRow).ClearContents
            
            lastRow = workingsSheet.Cells(workingsSheet.Rows.Count, "F").End(xlUp).row
            workingsSheet.Range("F1:F" & lastRow).ClearContents
        End If
    End If

    Application.DisplayAlerts = True ' Turn prompts back on
End Sub
