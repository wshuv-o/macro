Option Explicit

Sub ListExcelFilesInSubfolders()
    Dim mainFolder As String
    Dim subFolder As String
    Dim fileName As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim fso As Object
    Dim folder As Object
    Dim subFldr As Object
    Dim file As Object
    
    ' Let user select a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Main Folder"
        If .Show <> -1 Then Exit Sub
        mainFolder = .SelectedItems(1)
    End With
    
    ' Setup output headers
    Cells.Clear
    Range("A1").Value = "Subfolder"
    Range("B1").Value = "Excel Files"
    rowNum = 2
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(mainFolder)
    
    ' Loop through each subfolder
    For Each subFldr In folder.SubFolders
        Cells(rowNum, 1).Value = subFldr.Name
        colNum = 2
        
        ' Loop through files in each subfolder
        For Each file In subFldr.Files
            If LCase(fso.GetExtensionName(file.Name)) Like "xls*" Then
                Cells(rowNum, colNum).Value = file.Name
                colNum = colNum + 1
            End If
        Next file
        
        rowNum = rowNum + 1
    Next subFldr
    
    Columns.AutoFit
    MsgBox "Done! All subfolders and Excel files listed.", vbInformation
End Sub


Sub HighlightYearsInRange()
    Dim rng As Range, cell As Range
    Dim text As String, yearMatch As Object, re As Object
    Dim yearValue As Long, colorIndex As Long
    
    ' Define range
    Set rng = Range("A1:Y115")
    
    ' Clear any previous formatting
    rng.Interior.colorIndex = xlNone
    rng.Font.colorIndex = xlAutomatic
    
    ' Create RegExp to find 4-digit numbers between 2000–2027
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(20[0-2][0-9]|2027)"
    
    ' Loop through each cell
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            text = CStr(cell.Value)
            
            If re.Test(text) Then
                Set yearMatch = re.Execute(text)
                
                ' Reset cell format before applying highlights
                cell.Characters.Font.colorIndex = xlAutomatic
                cell.Interior.colorIndex = xlNone
                
                Dim m As Object
                For Each m In yearMatch
                    yearValue = CLng(m.Value)
                    
                    ' Map year to colorIndex (rotate colors for each year)
                    colorIndex = ((yearValue - 2000) Mod 6) + 3 ' colors 3–8
                    
                    ' Highlight the year portion
                    cell.Characters(m.FirstIndex + 1, Len(m.Value)).Font.colorIndex = colorIndex
                    cell.Characters(m.FirstIndex + 1, Len(m.Value)).Font.Bold = True
                Next m
            Else
                ' No year found ? fill cell dark gray
                cell.Interior.Color = RGB(64, 64, 64)
                cell.Font.Color = RGB(255, 255, 255)
            End If
        End If
    Next cell
    
    MsgBox "Year highlighting complete!", vbInformation
End Sub
Option Explicit

Sub HighlightLoanReview()
    Dim rng As Range, cell As Range
    Dim text As String
    
    ' Define the range
    Set rng = Range("A1:Y115")
    
    ' Clear previous highlighting
    rng.Interior.colorIndex = xlNone
    
    ' Loop through cells
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            text = LCase(CStr(cell.Value))
            
            ' Check if "loan review" appears anywhere in the text
            If InStr(text, "loan review") > 0 Then
                cell.Interior.Color = RGB(0, 200, 0)  ' Green
                cell.Font.Color = vbWhite
            End If
        End If
    Next cell 
    
    MsgBox "Cells containing 'Loan Review' highlighted in green.", vbInformation
End Sub

Sub HighlightRatingModel()
    Dim rng As Range, cell As Range
    Dim text As String
    
    ' Define the range
    Set rng = Range("A1:Y115")
    
    ' Clear previous highlighting
    rng.Interior.colorIndex = xlNone
    
    ' Loop through cells
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            text = LCase(CStr(cell.Value))
            
            ' Check if "loan review" appears anywhere in the text
            If InStr(text, "rating") > 0 Then
                cell.Interior.Color = RGB(200, 0, 0)  ' Green
                cell.Font.Color = vbWhite
            End If
        End If
    Next cell
    
    MsgBox "Cells containing 'Loan Review' highlighted in green.", vbInformation
End Sub


Function GetLatestUWFFile(folderPath As String) As String
    Dim fso As Object, folder As Object, file As Object
    Dim re As Object, matches As Object, m As Object
    Dim latestYear As Long, yearValue As Long
    Dim latestFileName As String
    
    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        GetLatestUWFFile = "Folder not found"
        Exit Function
    End If
    Set folder = fso.GetFolder(folderPath)
    
    ' Regex to detect years 2000–2027
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(200[0-9]|201[0-9]|202[0-7])"
    
    latestYear = 0
    latestFileName = ""
    
    ' Loop through files
    For Each file In folder.Files
        If LCase(file.Name) Like "uwf*" Then
            If re.Test(file.Name) Then
                Set matches = re.Execute(file.Name)
                For Each m In matches
                    yearValue = CLng(m.Value)
                    If yearValue > latestYear Then
                        latestYear = yearValue
                        latestFileName = file.Name
                    End If
                Next m
            End If
        End If
    Next file
    
    ' Return result
    If latestFileName <> "" Then
        GetLatestUWFFile = latestFileName
    Else
        GetLatestUWFFile = "No UWF file found"
    End If
End Function
Sub TestGetLatestUWFFile()
    Dim latestFile As String
    latestFile = GetLatestUWFFile("E:\Project Tapestry\Loan Review - 2025 - Batch 1\26490 - Colerain Hills Shopping Center (A1)")
    MsgBox "Latest UWF file: " & latestFile
End Sub

Sub WriteLatestUWFYear()
    Dim ws As Worksheet
    Dim rowNum As Long, colNum As Long
    Dim lastRow As Long
    Dim cellValue As String
    Dim yearValue As Long, latestYear As Long
    Dim re As Object, matches As Object
    
    ' Set worksheet and last row
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Regular expression to find years 2000–2027
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(20[0-2][0-9]|2027)"
    
    ' Loop through each row
    For rowNum = 1 To lastRow
        latestYear = 0
        
        ' Loop through columns B–Y
        For colNum = 2 To 25
            cellValue = Trim(CStr(ws.Cells(rowNum, colNum).Value))
            
            If LCase(cellValue) Like "uwf*" Then
                If re.Test(cellValue) Then
                    Set matches = re.Execute(cellValue)
                    Dim m As Object
                    For Each m In matches
                        yearValue = CLng(m.Value)
                        If yearValue > latestYear Then latestYear = yearValue
                    Next m
                End If
            End If
        Next colNum
        
        ' Write the latest year in column U if found
        If latestYear > 0 Then
            ws.Cells(rowNum, "U").Value = latestYear
        Else
            ws.Cells(rowNum, "U").Value = ""  ' blank if none found
        End If
    Next rowNum
    
    MsgBox "Latest UWF year written in column U.", vbInformation
End Sub


Sub MoveAndSortLRMToRight()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim vStartCol As Long: vStartCol = 22 ' column V
    Dim scanFirstCol As Long: scanFirstCol = 2 ' column B
    Dim scanLastCol As Long: scanLastCol = 25 ' column Y
    Dim outClearLastCol As Long: outClearLastCol = 52 ' column AZ - clear output area V:AZ first
    
    Dim re As Object, matches As Object
    Dim cellVal As String
    Dim tempFiles() As String
    Dim tempYears() As Long
    Dim cnt As Long
    Dim i As Long, j As Long
    Dim tFile As String, tYear As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' RegExp pattern to match years 2000-2027
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(200[0-9]|201[0-9]|202[0-7])"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Clear previous outputs in V:AZ (adjust if you need a larger area)
    ws.Range(ws.Cells(1, vStartCol), ws.Cells(lastRow, outClearLastCol)).ClearContents
    
    For r = 1 To lastRow
        cnt = 0
        
        ' First pass: collect LRM values and their years
        For c = scanFirstCol To scanLastCol
            cellVal = Trim(CStr(ws.Cells(r, c).Value))
            If Len(cellVal) > 0 Then
                If LCase(cellVal) Like "uwf*" Then
                    ' Found an LRM file. Extract year (if any)
                    If re.Test(cellVal) Then
                        Set matches = re.Execute(cellVal)
                        ' take the first matched year (if multiple, take the first occurrence)
                        tYear = CLng(matches(0).Value)
                    Else
                        tYear = -1 ' no valid year, put these after dated files
                    End If
                    
                    ' Grow arrays and store
                    cnt = cnt + 1
                    ReDim Preserve tempFiles(1 To cnt)
                    ReDim Preserve tempYears(1 To cnt)
                    tempFiles(cnt) = cellVal
                    tempYears(cnt) = tYear
                    
                    ' Clear original LRM cell to avoid duplicates when writing to V+
                    ws.Cells(r, c).ClearContents
                End If
            End If
        Next c
        
        ' If we found any LRM files, sort them by year (desc), then by filename (asc)
        If cnt > 1 Then
            For i = 1 To cnt - 1
                For j = i + 1 To cnt
                    ' Compare: primary by year (higher first). Note: year = -1 goes to the end.
                    If tempYears(i) < tempYears(j) Then
                        ' swap
                        tFile = tempFiles(i): tempFiles(i) = tempFiles(j): tempFiles(j) = tFile
                        tYear = tempYears(i): tempYears(i) = tempYears(j): tempYears(j) = tYear
                    ElseIf tempYears(i) = tempYears(j) Then
                        ' tie — sort by filename ascending (case-insensitive)
                        If StrComp(tempFiles(i), tempFiles(j), vbTextCompare) > 0 Then
                            tFile = tempFiles(i): tempFiles(i) = tempFiles(j): tempFiles(j) = tFile
                            tYear = tempYears(i): tempYears(i) = tempYears(j): tempYears(j) = tYear
                        End If
                    End If
                Next j
            Next i
        End If
        
        ' Write sorted LRM files starting at column V (vStartCol)
        If cnt > 0 Then
            For i = 1 To cnt
                ws.Cells(r, vStartCol + i - 1).Value = tempFiles(i)
            Next i
        End If
        
        ' Reset arrays for next row
        Erase tempFiles
        Erase tempYears
    Next r
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "LRM files sorted by year and moved to columns starting at V.", vbInformation
End Sub

