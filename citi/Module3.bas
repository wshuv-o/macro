Attribute VB_Name = "Module3"
Sub Tracker()
    '-----------------Tracker pull Vars-------------------
    Dim folderPath As String
    Dim subFolder As Object
    Dim file As Object
    Dim FileName As String
    Dim wb As Workbook
    Dim CashFlowSheet As Worksheet
    Dim loanSummaryRow As Long
    Dim destSheet As Worksheet
    Dim incrementRow As Integer
    Dim increment As Integer
    Dim subFolderName As String
    Dim subFolderPart1 As String
    Dim subFolderPart2 As String
    Dim spacePos As Integer
    Dim fso As Object
    
    ' Set destination sheet for output
    Set destSheet = ThisWorkbook.Sheets("Tracker")
    incrementRow = 2
    
    ' Select folder containing subfolders/files.xlsm
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Subfolders"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable screen updating, calculations, and events to improve performance
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Loop through all subfolders in the selected folder
    For Each subFolder In fso.GetFolder(folderPath).SubFolders
        subFolderName = subFolder.Name
        spacePos = InStr(subFolderName, "-") ' Find the position of the first space

        If spacePos > 0 Then
            ' Split the subfolder name into two parts (before and after the first space)
            subFolderPart1 = Left(subFolderName, spacePos - 1)
            subFolderPart2 = Mid(subFolderName, spacePos + 1)
            
            ' Reset increment for each new subfolder
            increment = 1
            
            ' Loop through each file in the subfolder
            For Each file In subFolder.Files
                FileName = file.Name

                ' Check if the file starts with 'UW' and is an Excel file
                If FileName Like "UW*" And _
                   (Right(FileName, 4) = ".xls" Or Right(FileName, 5) = ".xlsx" Or Right(FileName, 5) = ".xlsm") Then

                    ' Open the workbook for Tracker data
                    Set wb = Workbooks.Open(file.path, ReadOnly:=True)
                    
                    ' Check if "Cash Flow" sheet exists
                    Set CashFlowSheet = Nothing
                    On Error Resume Next
                    Set CashFlowSheet = wb.Sheets("Cash Flow")
                    On Error GoTo 0
                    
                    ' Process Tracker Details if the sheet exists
                    If Not CashFlowSheet Is Nothing Then
                        loanSummaryRow = 66
                        
                        ' Populate Tracker data columns A to G
                        destSheet.Cells(incrementRow, 1).Value = Trim(subFolderPart1)                                             ' Loan ID
                        destSheet.Cells(incrementRow, 2).Value = Trim(subFolderPart1) & "-" & increment                           ' Asset ID
                        destSheet.Cells(incrementRow, 3).Value = Trim(subFolderPart2)                                             ' Loan Name
                        destSheet.Cells(incrementRow, 4).Value = CashFlowSheet.Cells(6, 5).Value                           ' Asset name
                        destSheet.Cells(incrementRow, 5).Value = CashFlowSheet.Cells(7, 5).Value & ", " & CashFlowSheet.Cells(7, 7).Value                       ' Address
                        destSheet.Cells(incrementRow, 6).Value = CashFlowSheet.Cells(8, 5).Value                            ' Loan Summary
                        destSheet.Cells(incrementRow, 7).Value = subFolder.Name                                             ' Loan Name from Folder
                        destSheet.Cells(incrementRow, 8).Value = fso.GetFolder(folderPath).Name                                           ' Batch
                        destSheet.Cells(incrementRow, 9).Value = "=OFFSET(Mapping!$C$4, MATCH(F" & incrementRow & ", Mapping!$B$5:$B$60, 0), 0)"
                        
                        
                        ' Increment row for the next file
                        incrementRow = incrementRow + 1
                        ' Increment Asset ID counter for the next UW file in the same subfolder
                        increment = increment + 1
                    Else
                        ' Log missing Cash Flow sheet
                        Debug.Print "Cash Flow sheet not found in " & FileName
                    End If
                    
                    wb.Close False
                End If
            Next file
        End If
    Next subFolder

    ' Restore application settings
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Data extraction complete!", vbInformation
End Sub
Function GetColumnNumber(ws As Worksheet, searchValue As String, rowNumber As Long) As Long
    Dim rng As Range
    
    ' Search for the string in the specified row
    Set rng = ws.Rows(rowNumber).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Return column number if found, else return 0
    If Not rng Is Nothing Then
        GetColumnNumber = rng.Column
    Else
        GetColumnNumber = 5
    End If
End Function
Sub TrackerV1()
    '-----------------Tracker pull Vars-------------------
    Dim folderPath As String
    Dim subFolder As Object
    Dim file As Object
    Dim FileName As String
    Dim wb As Workbook
    Dim CashFlowSheet As Worksheet
    Dim ALASheet As Worksheet
    Dim loanSummaryRow As Long
    Dim destSheet As Worksheet
    Dim incrementRow As Integer
    Dim increment As Integer
    Dim subFolderName As String
    Dim subFolderPart1 As String
    Dim subFolderPart2 As String
    Dim spacePos As Integer
    Dim fso As Object
    
    ' Set destination sheet for output
    Set destSheet = ThisWorkbook.Sheets("Tracker")
    incrementRow = 2
    
    ' Select folder containing subfolders/files.xlsm
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Subfolders"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable screen updating, calculations, and events to improve performance
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Loop through all subfolders in the selected folder
    For Each subFolder In fso.GetFolder(folderPath).SubFolders
        subFolderName = subFolder.Name
        spacePos = InStr(subFolderName, "-") ' Find the position of the first hyphen

        If spacePos > 0 Then
            ' Split the subfolder name into two parts (before and after the first hyphen)
            subFolderPart1 = Trim(Left(subFolderName, spacePos - 1))
            subFolderPart2 = Trim(Mid(subFolderName, spacePos + 1))
            
            ' Reset increment for each new subfolder
            increment = 1
            
            ' Loop through each file in the subfolder
            For Each file In subFolder.Files
                FileName = file.Name

                ' Check if the file starts with 'UW' and is an Excel file
                If FileName Like "UW*" And _
                   (Right(FileName, 4) = ".xls" Or Right(FileName, 5) = ".xlsx" Or Right(FileName, 5) = ".xlsm") Then

                    ' Open the workbook for Tracker data
                    Set wb = Workbooks.Open(file.path, ReadOnly:=True)
                    Dim si, property, address, city, state, zip As Variant
                    ' Check if "ALA" sheet exists
                    Set ALASheet = Nothing
                    On Error Resume Next
                    Set ALASheet = wb.Sheets("Summary")
                    si = 2
                    property = GetColumnNumber(ALASheet, "Property", 3)
                    address = GetColumnNumber(ALASheet, "Address", 3)
                    city = GetColumnNumber(ALASheet, "City", 3)
                    state = GetColumnNumber(ALASheet, "State", 3)
                    zip = GetColumnNumber(ALASheet, "Zip", 3)
                    
                    On Error GoTo 0
                    
                    Set CashFlowSheet = Nothing
                    On Error Resume Next
                    Set CashFlowSheet = wb.Sheets("Cash Flow")
                    On Error GoTo 0
                    
                    
                    
                    
                    ' If ALA sheet exists, iterate through rows until empty
                    If Not ALASheet Is Nothing Then
                        loanSummaryRow = 5
                        
                        ' Iterate until an empty row is found in column C (Asset Name)
                        While Not IsEmpty(ALASheet.Cells(loanSummaryRow, 3).Value)
                            ' Populate Tracker data columns A to I
                            destSheet.Cells(incrementRow, 1).Value = subFolderPart1                                             ' Loan ID
                            destSheet.Cells(incrementRow, 2).Value = subFolderPart1 & "-" & ALASheet.Cells(loanSummaryRow, si).Value                           ' Asset ID
                            destSheet.Cells(incrementRow, 3).Value = subFolderPart2                                             ' Loan Name
                            destSheet.Cells(incrementRow, 4).Value = ALASheet.Cells(loanSummaryRow, property).Value                    ' Asset name
                            destSheet.Cells(incrementRow, 5).Value = ALASheet.Cells(loanSummaryRow, address).Value & ", " & _
                                                                    ALASheet.Cells(loanSummaryRow, city).Value & ", " & _
                                                                    ALASheet.Cells(loanSummaryRow, state).Value & ", " & ALASheet.Cells(loanSummaryRow, zip).Value           ' Address
                            destSheet.Cells(incrementRow, 6).Value = CashFlowSheet.Cells(6, 4).Value                                ' Loan Summary
                            destSheet.Cells(incrementRow, 7).Value = subFolder.Name                                             ' Loan Name from Folder
                            destSheet.Cells(incrementRow, 8).Value = fso.GetFolder(folderPath).Name                             ' Batch
                            destSheet.Cells(incrementRow, 9).Value = "=OFFSET(Mapping!$C$4, MATCH(F" & incrementRow & ", Mapping!$B$5:$B$60, 0), 0)"
                            
                            ' Increment row for the next entry
                            incrementRow = incrementRow + 1
                            loanSummaryRow = loanSummaryRow + 1
                        Wend
                        
                        ' Increment Asset ID counter after processing all rows in the ALA sheet
                        increment = increment + 1
                    Else
                        ' Process Tracker Details if Cash Flow sheet exists
                        If Not CashFlowSheet Is Nothing Then
                            loanSummaryRow = 66
                            
                            ' Populate Tracker data columns A to I from Cash Flow sheet (single row)
                            destSheet.Cells(incrementRow, 1).Value = subFolderPart1                                             ' Loan ID
                            destSheet.Cells(incrementRow, 2).Value = subFolderPart1 & "-" & increment                           ' Asset ID
                            destSheet.Cells(incrementRow, 3).Value = subFolderPart2                                             ' Loan Name
                            destSheet.Cells(incrementRow, 4).Value = CashFlowSheet.Cells(6, 5).Value                           ' Asset name
                            destSheet.Cells(incrementRow, 5).Value = CashFlowSheet.Cells(7, 5).Value & ", " & _
                                                                    CashFlowSheet.Cells(7, 7).Value                    ' Address
                            destSheet.Cells(incrementRow, 6).Value = CashFlowSheet.Cells(8, 5).Value                           ' Loan Summary
                            destSheet.Cells(incrementRow, 7).Value = subFolder.Name                                             ' Loan Name from Folder
                            destSheet.Cells(incrementRow, 8).Value = fso.GetFolder(folderPath).Name                             ' Batch
                            destSheet.Cells(incrementRow, 9).Value = "=OFFSET(Mapping!$C$4, MATCH(F" & incrementRow & ", Mapping!$B$5:$B$60, 0), 0)"
                            
                            ' Increment row for the next file
                            incrementRow = incrementRow + 1
                            ' Increment Asset ID counter for the next UW file
                            increment = increment + 1
                        Else
                            ' Log missing ALA and Cash Flow sheets
                            Debug.Print "Neither ALA nor Cash Flow sheet found in " & FileName
                        End If
                    End If
                    
                    wb.Close False
                End If
            Next file
        End If
    Next subFolder

    ' Restore application settings
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Data extraction complete!", vbInformation
End Sub

Function FindAdjacentValueWS(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim r As Long, c As Long
    Dim foundCell As Range
    Dim checkCell As Range
    
    ' Search for the searchText in the given range
'    For Each cell In searchRange
 '       If cell.Value = searchText Then
  '          Set foundCell = cell
     '       Exit For
   '     End If
    'Next cell
    For Each cell In searchRange
    If Not IsError(cell.Value) Then
        If Trim(cell.Value) = searchText Then
            Set foundCell = cell
            Exit For
        End If
    End If
Next cell

    If foundCell Is Nothing Then
        FindAdjacentValue = "Not Found"
        Exit Function
    End If
    
    r = foundCell.Row
    c = foundCell.Column
    
    If direction = "right" Then
        Dim i As Integer
        For i = 1 To maxRight
            On Error Resume Next
            Set checkCell = ws.Cells(r, c + i)
            If Not checkCell.MergeCells Then
                If Trim(checkCell.Value) <> "" Then
                    FindAdjacentValue = checkCell.Value
                    Exit Function
                End If
            Else
                Set checkCell = checkCell.MergeArea.Cells(1, 1)
                If Trim(checkCell.Value) <> "" Then
                    FindAdjacentValue = checkCell.Value
                    Exit Function
                End If
            End If
            On Error GoTo 0
        Next i
        FindAdjacentValue = "No Value Found"
        
    ElseIf direction = "down" Then
        Dim j As Integer
        For j = 1 To maxDown
            Set checkCell = ws.Cells(r + j, c)
            If Trim(checkCell.Value) <> "" Then
                FindAdjacentValue = checkCell.Value
                Exit Function
            End If
        Next j
        FindAdjacentValue = "No Value Found"
    Else
        FindAdjacentValue = "Invalid Direction"
    End If
End Function

Function getValueAtIntersection(ws As Worksheet) As Variant
    Dim xCell As Range, yCell As Range
    Dim resultCell As Range

    Dim xHeader As String
    Dim yHeader As String
    xHeader = "Underwritten"
    yHeader = "Debt Service on Recommended loan"

    ' Search for xHeader in rows 20 to 30, all columns
    Set xCell = ws.Range("A20:AP30").Find(What:=xHeader, LookIn:=xlValues, LookAt:=xlWhole)
    
    If xCell Is Nothing Then
        MsgBox "X Header Not Found"
        Exit Function
    End If

    ' Search for yHeader in columns A to E, all rows
    Set yCell = ws.Range("A1:E100").Find(What:=yHeader, LookIn:=xlValues, LookAt:=xlWhole)
    If yCell Is Nothing Then
        MsgBox "Y Header Not Found"
        Exit Function
    End If

    ' Return the intersecting cell's value
    Set resultCell = ws.Cells(yCell.Row, xCell.Column)
    getValueAtIntersection = resultCell.Value
End Function



Sub Loan()
    '-----------------Tracker pull Vars-------------------
    Dim CashFlowSheet As Worksheet
    Dim ALASheet As Worksheet
    Dim loanSummarySheet As Worksheet
    Dim loanSummaryRow As Long
    Dim destSheet As Worksheet
    Dim incrementRow As Integer
    Dim increment As Integer
    Dim wb As Workbook
    Dim si, propertyCol, addressCol, cityCol, stateCol, zipCol As Variant
    
    Set wb = ThisWorkbook
    Set destSheet = wb.Sheets("Loan")
    incrementRow = 6
    increment = 1

    ' Disable screen updating, calculations, and events to improve performance
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Loop through all worksheets from the 12th to the last
    Dim i As Integer
    For i = 12 To wb.Sheets.count
        Set CashFlowSheet = wb.Sheets(i)

        ' Assuming loanSummarySheet = CashFlowSheet or another known name
        Set loanSummarySheet = CashFlowSheet ' Adjust as needed

        destSheet.Cells(incrementRow, 2).Value = "=TEXTJOIN("", "", TRUE, FILTER(Tracker!B:B, Tracker!A:A=A" & incrementRow & "))"
        destSheet.Cells(incrementRow, 3).Value = ""
        destSheet.Cells(incrementRow, 4).Value = CashFlowSheet.Cells(6, 23).Value
        destSheet.Cells(incrementRow, 5).Value = CashFlowSheet.Cells(6, 23).Value
        destSheet.Cells(incrementRow, 6).Value = "=EOMONTH(C" & incrementRow & ",AH" & incrementRow & ")+1"
        destSheet.Cells(incrementRow, 7).Value = getValueAtIntersection(CashFlowSheet)
        destSheet.Cells(incrementRow, 8).Value = "=G" & incrementRow & "/12"
        destSheet.Cells(incrementRow, 9).Value = loanSummarySheet.Cells(12, 23).Value
        destSheet.Cells(incrementRow, 10).Value = loanSummarySheet.Cells(12, 20).Value
        destSheet.Cells(incrementRow, 11).Value = ""
        destSheet.Cells(incrementRow, 12).Value = ""
        destSheet.Cells(incrementRow, 13).Value = "=EOMONTH(C" & incrementRow & ",N" & incrementRow & ")+1"
'        destSheet.Cells(incrementRow, 14).Value = loanSummarySheet.Range("IOPeriods").Value
 '       destSheet.Cells(incrementRow, 15).Value = loanSummarySheet.Range("LoanAnalysis_Rate").Value
  '      destSheet.Cells(incrementRow, 16).Value = loanSummarySheet.Range("LS_IndexType").Value
   '     destSheet.Cells(incrementRow, 17).Value = loanSummarySheet.Range("LS_Spread").Value
    '    destSheet.Cells(incrementRow, 18).Value = ""
     '   destSheet.Cells(incrementRow, 19).Value = ""
      '  destSheet.Cells(incrementRow, 20).Value = ""
       ' destSheet.Cells(incrementRow, 21).Value = ""
        'destSheet.Cells(incrementRow, 22).Value = loanSummarySheet.Range("W12").Value
        'destSheet.Cells(incrementRow, 23).Value = loanSummarySheet.Range("R13").Value
        'destSheet.Cells(incrementRow, 24).Value = loanSummarySheet.Range("AmortTerm").Value

        incrementRow = incrementRow + 1
        increment = increment + 1
    Next i

    ' Restore application settings
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Data extraction complete!", vbInformation
End Sub

