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
    Dim targetRange As Range
    Dim copiedRange As Range

    ' Disable pop-up alerts by default
    Application.DisplayAlerts = False

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
        sheetCounter = 0
        Set sourceWorkbook = Workbooks.Open(folderPath & fileName, ReadOnly:=True)

        For Each sheet In sourceWorkbook.Sheets
            ' Check if the sheet name contains "Cash Flow" but does not contain "Details" or "Footnote"
            If sheet.Name Like "*Cash Flow*" And Not sheet.Name Like "*Aggregate Cash Flow*" And Not sheet.Name Like "*Cash Flow Detail*" And Not sheet.Name Like "*Cash Flow Footnote*" Then

                ' Define the range to copy
                Set copiedRange = sheet.Range("H16:AG89")

                ' Copy the content (H5:AG89) from the source sheet
                copiedRange.Copy

                ' Generate a unique sheet name based on the sheet name
                newSheetName = sheet.Range("H5").value

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

                ' If the sheet already exists, append a counter until a unique name is found
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
                Set newSheet = ThisWorkbook.Sheets.Add
                newSheet.Name = newSheetName

                ' Explicitly define the target range size (same as the copied range)
                Set targetRange = newSheet.Range("A1").Resize(copiedRange.Rows.Count, copiedRange.Columns.Count)

                ' Paste the copied values and formatting into the new sheet in ThisWorkbook
                targetRange.PasteSpecial Paste:=xlPasteValues
                targetRange.PasteSpecial Paste:=xlPasteFormats

                ' Find the row in the "H" column that contains "Net Cash Flow"
                Set netCashFlowCell = Nothing
                For Each cell In newSheet.Range("H:H")
                    If cell.value Like "*Net Cash Flow*" Then
                        Set netCashFlowCell = cell
                        Exit For
                    End If
                Next cell

                ' If a cell is found, paste the copied content at the corresponding row
                If Not netCashFlowCell Is Nothing Then
                    ' Define the target range (row where "Net Cash Flow" was found)
                    Set targetRange = netCashFlowCell.Offset(0, 0).Resize(1, 30) ' H to AG (30 columns in total)

                    ' Paste values
                    targetRange.PasteSpecial Paste:=xlPasteValues
                    ' Paste formatting
                    targetRange.PasteSpecial Paste:=xlPasteFormats

                    ' Clear the clipboard
                    Application.CutCopyMode = False
                End If
            End If
        Next sheet

        sourceWorkbook.Close False
        fileName = Dir
    Loop

    ' Re-enable pop-up alerts
    Application.DisplayAlerts = True

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
    Dim targetRange As Range
    Dim copiedRange As Range
    Dim lastRow As Long
    Dim startRow As Long

    ' Disable pop-up alerts by default
    Application.DisplayAlerts = False

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

                    ' Copy the content from the source sheet
                    copiedRange.Copy

                    ' Generate a unique sheet name based on the sheet name
                    newSheetName = sheet.Range("H5").value

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

                    ' If the sheet already exists, append a counter until a unique name is found
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
                    Set newSheet = ThisWorkbook.Sheets.Add
                    newSheet.Name = newSheetName

                    ' Explicitly define the target range size (same as the copied range)
                    Set targetRange = newSheet.Range("A1").Resize(copiedRange.Rows.Count, copiedRange.Columns.Count)

                    ' Paste the copied values and formatting into the new sheet in ThisWorkbook
                    targetRange.PasteSpecial Paste:=xlPasteValues
                    targetRange.PasteSpecial Paste:=xlPasteFormats

                    ' Clear the clipboard
                    Application.CutCopyMode = False
                End If
            End If
        Next sheet

        sourceWorkbook.Close False
        fileName = Dir
    Loop

    ' Re-enable pop-up alerts
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True
    MsgBox "Cash Flow sheets extracted and renamed successfully!", vbInformation
End Sub




faster-------------------
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
    Dim targetRange As Range
    Dim copiedRange As Range
    Dim lastRow As Long
    Dim startRow As Long

    ' Disable pop-up alerts, screen updating, and calculations for faster processing
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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
    
    Do While fileName <> ""
        sheetCounter = 0
        Set sourceWorkbook = Workbooks.Open(folderPath & fileName, ReadOnly:=True)

        For Each sheet In sourceWorkbook.Sheets
            ' Check if the sheet name contains "Cash Flow" but does not contain "Details" or "Footnote"
            If sheet.Name Like "*Cash Flow*" And Not sheet.Name Like "*Aggregate Cash Flow*" And Not sheet.Name Like "*Cash Flow Detail*" And Not sheet.Name Like "*Cash Flow Footnote*" Then

                ' Find the first occurrence of "Net Cash Flow" in column H
                Set netCashFlowCell = Nothing
                For Each cell In sheet.Range("H16:H" & sheet.Cells(sheet.Rows.Count, "H").End(xlUp).Row)
                    If cell.Value Like "*Net Cash Flow*" Then
                        Set netCashFlowCell = cell
                        Exit For
                    End If
                Next cell

                ' If we found "Net Cash Flow", set the range dynamically
                If Not netCashFlowCell Is Nothing Then
                    ' Set the range to start at H16 and end at the row where "Net Cash Flow" was found
                    Set copiedRange = sheet.Range("H16:AG" & netCashFlowCell.Row)

                    ' Generate a unique sheet name based on the sheet name
                    newSheetName = sheet.Range("H5").Value

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

                    ' If the sheet already exists, append a counter until a unique name is found
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
                    Set newSheet = ThisWorkbook.Sheets.Add
                    newSheet.Name = newSheetName

                    ' Explicitly define the target range size (same as the copied range)
                    Set targetRange = newSheet.Range("A1").Resize(copiedRange.Rows.Count, copiedRange.Columns.Count)

                    ' Copy values directly (no clipboard used)
                    targetRange.Value = copiedRange.Value

                    ' Copy formats directly (no clipboard used)
                    copiedRange.Copy
                    targetRange.PasteSpecial Paste:=xlPasteFormats

                    ' Clear the clipboard
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



faster2--------------------------------
