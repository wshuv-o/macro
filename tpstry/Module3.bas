Attribute VB_Name = "Module3"
Sub Reset()
    Dim frReset As Object
    Set frReset = New UserForm2
    frReset.Show
End Sub

Sub PullRentRoll()

    Dim sourceWorkbook As Workbook
    Dim sheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim invalidChars As String
    Dim i As Integer
    Dim sheetCounter As Integer
    Dim folderPath As String
    Dim FileName As String
    Dim netRentRollCell As Range
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
    FileName = Dir(folderPath & "*.xlsm")
    
    Do While FileName <> ""
        sheetCounter = 0
        Set sourceWorkbook = Workbooks.Open(folderPath & FileName, ReadOnly:=True)

        For Each sheet In sourceWorkbook.Sheets
            ' Check if the sheet name contains "Rent Roll" but does not contain "Details" or "Footnote"
            If sheet.Name Like "*Rent Roll*" And Not sheet.Name Like "*Rent Roll Analytics*" And Not sheet.Name Like "*Aggregate Rent Roll*" And Not sheet.Name Like "*Rent Roll Footnote*" Then

                ' Find the first occurrence of "Total" in column E
                Set netRentRollCell = Nothing
                For Each cell In sheet.Range("E3:E" & sheet.Cells(sheet.Rows.Count, "E").End(xlUp).Row)
                    If cell.value Like "*Total" Then
                        Set netRentRollCell = cell
                        Exit For
                    End If
                Next cell

                ' If we found "Net Rent Roll", set the range dynamically
                If Not netRentRollCell Is Nothing Then
                    Set copiedRange = sheet.Range("E3:AN" & netRentRollCell.Row)                        ' Set the range to start at H16 and end at the row where "Net Rent Roll" was found
                    dataArray = copiedRange.value                                                       ' Load data from copiedRange into an array for faster manipulation
                    newSheetName = sheet.Range("E4").value                                              ' Generate a unique sheet name based on the sheet name
                    invalidChars = "/\?*:[]"                                                            ' Remove invalid characters from the sheet name
                    For i = 1 To Len(invalidChars)
                        newSheetName = Replace(newSheetName, Mid(invalidChars, i, 1), "")
                    Next i

                    If Len(newSheetName) > 23 Then                                                      ' Trim name length and append counter for uniqueness
                        newSheetName = Left(newSheetName, 23)
                    End If
                    
                    Dim tempSheetName As String                                                         ' Check if a sheet with this name already exists, and if so, append a counter
                    tempSheetName = newSheetName
                    sheetCounter = 0                                                                    ' Reset the counter each time

                    On Error Resume Next
                    Set newSheet = ThisWorkbook.Sheets(tempSheetName)
                    On Error GoTo 0

                     ''If the sheet already exists, append a counter until a unique name is found
                    While Not newSheet Is Nothing
                        sheetCounter = sheetCounter + 1
                        tempSheetName = "RR " & newSheetName & " (" & sheetCounter & ")"
                        Set newSheet = Nothing
                       On Error Resume Next
                        Set newSheet = ThisWorkbook.Sheets(tempSheetName)
                        On Error GoTo 0
                    Wend

                    
                    newSheetName = tempSheetName                                                                    ' Now, set the unique sheet name
                    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))   ' Create a new sheet in ThisWorkbook with the unique name
                    newSheet.Name = newSheetName
                    newSheet.Range("A1").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray       ' Write the dataArray back to the new sheet in one go
                    copiedRange.Copy                                                                                ' Copy formats from the original range
                    newSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
                End If
            End If
        Next sheet

        sourceWorkbook.Close False
        FileName = Dir
    Loop

    ' Re-enable pop-up alerts, screen updating, and calculations
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Rent Roll sheets extracted and renamed successfully!", vbInformation
End Sub

