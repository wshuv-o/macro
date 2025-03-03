' Older Version
Sub ExtractCashFlowSheetIGNORE()
    Dim sourceWorkbook As Workbook
    Dim sheet As Worksheet

    ' Select the source .xlsm file
    With Application.FileDialog(msoFileDialogFilePicker)
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
    With Application.FileDialog(msoFileDialogFolderPicker)
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
                newSheetName = newSheet.Range("H5").Value

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

