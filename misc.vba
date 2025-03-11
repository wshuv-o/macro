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
' Older Version
Sub ExtractCashFlowSheetIGNORE()
    Dim sourceWorkbook As Workbook
    Dim sheet As Worksheet

    ' Select the source .xlsm file
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Select the .xlsm File"
        .Filters.Add "Excel Files", "*.xlsm"
        If .Show = -1 Then
            Set sourceWorkbook = Workbooks.Open(.SelectedItems(1), ReadOnly:=True)
        Else
            MsgBox "No file selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    Application.ScreenUpdating = False
    For Each sheet In sourceWorkbook.Sheets
        If sheet.Name Like "*Cash Flow" Then sheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Next sheet

    sourceWorkbook.Close False
    Application.ScreenUpdating = True
    MsgBox "Cash Flow sheets extracted and appended successfully!", vbInformation
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


Sub ExtractCashFlowSheetsTEST()
    Dim sourceWorkbook As Workbook
    Dim sheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim invalidChars As String
    Dim i As Integer
    Dim sheetCounter As Integer
    Dim folderPath As String
    Dim fileName As String
    Dim frm As UserForm1
    Dim fileCount As Integer
    Set frm = UserForm1
    
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

    ' Count the number of .xlsm files in the folder
    fileName = Dir(folderPath & "*.xlsm")


    fileCount = 0
    Do While fileName <> ""
        fileCount = fileCount + 1
        fileName = Dir ' Get the next file
    Loop

    ' Initialize the form with the number of files
    frm.InitializeForm fileCount ' Initialize form with file count
    frm.Show

    ' Loop through all the .xlsm files in the selected folder
    fileName = Dir(folderPath & "*.xlsm") ' Get the first file again
    Application.ScreenUpdating = False

    Do While fileName <> ""
        sheetCounter = 1
        Set sourceWorkbook = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        
        For Each sheet In sourceWorkbook.Sheets
            ' Check if the sheet name contains "Cash Flow" but does not contain "Details" or "Footnote"
            If sheet.Name Like "*Cash Flow*" And Not sheet.Name Like "*Aggregate Cash Flow*" _
                And Not sheet.Name Like "*Cash Flow Detail*" And Not sheet.Name Like "*Cash Flow Footnote*" Then

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
        
        ' Increment the progress and update the form after processing the current file
        'frm.IncrementCount ' Increment count after each file is processed
        'frm.UpdateProgress ' Update progress bar
        fileName = Dir ' Get the next file
    Loop

    Application.ScreenUpdating = True
    MsgBox "Cash Flow sheets extracted and renamed successfully!", vbInformation
End Sub






https://docs.google.com/document/d/1S_hpgkH_eEWYFZ1zDBNHAH9F7W9Jv_1x23BBwM3c6IY/edit?tab=t.0





Sub PullTrackerDetails()

    Dim folderPath As String
    Dim subFolder As Object
    Dim file As Object
    Dim fileName As String
    Dim wb As Workbook
    Dim loanSummarySheet As Worksheet
    Dim loanSummaryRow As Long
    Dim destSheet As Worksheet
    Dim incrementRow As Integer
    Dim increment As Integer
    Dim subFolderName As String
    Dim subFolderPart1 As String
    Dim subFolderPart2 As String
    Dim spacePos As Integer
    Dim fso As Object ' FileSystemObject for folder and file iteration

    ' Set the destination sheet for output
    Set destSheet = ThisWorkbook.Sheets("Tracker")

    ' Initialize increment
  
    incrementRow = 2

    ' Ask the user to select a folder containing subfolders
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Subfolders"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False

    ' Loop through each subfolder in the selected folder
    For Each subFolder In fso.GetFolder(folderPath).SubFolders

        subFolderName = subFolder.Name
        spacePos = InStr(subFolderName, " ") ' Find the position of the first space

        If spacePos > 0 Then
            ' Split the subfolder name into two parts (before and after the first space)
            subFolderPart1 = Left(subFolderName, spacePos - 1)
            subFolderPart2 = Mid(subFolderName, spacePos + 1)
            
            For Each file In subFolder.Files
                fileName = file.Name

                ' Check if the file starts with 'UW' and is an Excel file (.xls, .xlsx, .xlsm)
                If fileName Like "UW*" And _
                   (Right(fileName, 4) = ".xls" Or Right(fileName, 5) = ".xlsx" Or Right(fileName, 5) = ".xlsm") Then
                   

                    Set wb = Workbooks.Open(file.Path, ReadOnly:=True)
                    Set loanSummarySheet = wb.Sheets("Loan Analysis")

                    If Not loanSummarySheet Is Nothing Then
                        ' Set the row to start from (66th row)
                        loanSummaryRow = 66
                        increment = 1 ' Start from row 2 to avoid overwriting header
                        ' Loop through rows starting from 66 and onward
                        Do While Not IsEmpty(loanSummarySheet.Cells(loanSummaryRow, 6).Value) And _
                            Not loanSummarySheet.Cells(loanSummaryRow, 6).Value Like "*Total*"

                            ' Populate columns A, B, C, and G (already done earlier in the code)
                                        ' Store the left part in A column (subfolder name before the space)
                            destSheet.Cells(incrementRow, 1).Value = subFolderPart1
                
                            ' In B column, store the left part followed by increment (e.g., A2 & "-1")
                            destSheet.Cells(incrementRow, 2).Value = subFolderPart1 & "-" & (increment)
                
                            ' In C column, store the right part of the split (after the first space)
                            destSheet.Cells(incrementRow, 3).Value = subFolderPart2
                
                            ' Set subfolder name in G column
                            destSheet.Cells(incrementRow, 7).Value = subFolder.Name

                            ' D - Loan Summary sheet's column 6 (F)
                            destSheet.Cells(incrementRow, 4).Value = loanSummarySheet.Cells(loanSummaryRow, 6).Value ' D

                            ' E - Loan Summary sheet's columns 20 to 24 (T66:X66)
                            destSheet.Cells(incrementRow, 5).Value = Join(Application.Transpose( _
                                Application.Transpose(loanSummarySheet.Range("T" & loanSummaryRow & ":X" & loanSummaryRow).Value)), ", ") ' E

                            ' F - Loan Summary sheet's column 9 (I)
                            destSheet.Cells(incrementRow, 6).Value = loanSummarySheet.Cells(loanSummaryRow, 9).Value ' F

                            ' Increment for the next row
                            increment = increment + 1
                            incrementRow = incrementRow + 1
                            loanSummaryRow = loanSummaryRow + 1
                        Loop
                    End If

                    ' Close the workbook without saving
                    wb.Close False

                End If
            Next file
        End If

        ' Move to the next subfolder
        incrementRow = incrementRow + 1 ' Increment destination row

    Next subFolder

    ' Re-enable screen updating
    Application.ScreenUpdating = True

    ' Notify the user that the task is complete
    MsgBox "Data extraction complete!", vbInformation
    Exit Sub

FileError:
    MsgBox "Error opening file: " & file.Path, vbCritical
    Resume Next

End Sub


https://jobs.bdjobs.com/jobdetails.asp?id=1347145&ln=1







-----------------------------



Sub UnmergeHeader()
    Rows("1:5").Select
    Selection.UnMerge
    Range("D2").Select
    Selection.Cut
    Range("E1").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("J1").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("O1").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("T1").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("Y1").Select
    ActiveSheet.Paste
    Range("F4").Select
    Selection.Cut
    Range("E4").Select
    ActiveSheet.Paste
    Range("K4").Select
    Selection.Cut
    Range("J4").Select
    ActiveSheet.Paste
    Range("P4").Select
    Selection.Cut
    Range("O4").Select
    ActiveSheet.Paste
    Range("U4").Select
    Selection.Cut
    Range("T4").Select
    ActiveSheet.Paste
    Range("Z4").Select
    Selection.Cut
    Range("Y4").Select
    ActiveSheet.Paste
End Sub

Sub RemoveRowsWithInvalidStyle()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellStyle As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through each row starting from row 7
    For i = lastRow To 7 Step -1 ' Loop from bottom to top to avoid skipping rows after deletion
        ' Check if the cell style in column A is not "#_0_E"
        cellStyle = ws.Cells(i, 5).Style ' Assuming you're checking style in column A, you can adjust the column as needed
        If cellStyle <> "#_0_E" Then
            ws.Rows(i).Delete
        End If
    Next i

End Sub

Sub RemoveAmountColumn()

    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim header As String

    ' Set the current active sheet as the target worksheet
    Set ws = ActiveSheet

    ' Find the last used column in the first row (headers)
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through each column starting from column B
    For col = lastCol To 2 Step -1 ' Loop from right to left to avoid skipping columns after deletion
        ' Get the header value in the current column
        header = ws.Cells(1, col).value

        ' Check if the header contains the word "amount" (case-insensitive)
        If InStr(1, header, "Amount", vbTextCompare) > 0 Then
            ' If the header contains "amount", delete the entire column
            ws.Columns(col).Delete
        End If
    Next col

End Sub


Sub Action()
    UnmergeHeader
    RemoveRowsWithInvalidStyle
    RemoveAmountColumn
End Sub









=-----------signout



