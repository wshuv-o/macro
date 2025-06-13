Attribute VB_Name = "Module2"
Sub ImportCashflowSheetsFromUWFiles()
    Dim mainFolder As String
    Dim fso As Object, folder As Object, subfolderObj As Object
    Dim file As Object, wbSource As Workbook
    Dim ws As Worksheet, newWS As Worksheet
    Dim currentWB As Workbook
    Dim sheetCopied As Boolean
    Dim safeSheetName As String
    Dim namedRange As Name
    Dim targetRange As Range
    Dim usedRange As Range
    Dim col As Long

    ' Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Main Folder"
        If .Show <> -1 Then Exit Sub
        mainFolder = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(mainFolder)
    Set currentWB = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through each subfolder
    For Each subfolderObj In folder.SubFolders
        For Each file In subfolderObj.Files
            If file.Name Like "UW*.xls*" Then
                On Error Resume Next
                Set wbSource = Workbooks.Open(file.path, ReadOnly:=True)
                If Err.Number <> 0 Then
                    Debug.Print "Failed to open: " & file.path
                    Err.Clear
                    GoTo NextFile
                End If
                On Error GoTo 0
                
                For Each ws In wbSource.Worksheets
                    If LCase(ws.Name) Like "cash flow" Then
                                            
                        On Error Resume Next
                        Set namedRange = wbSource.Names("sizingcf")
                        On Error GoTo 0
                        
                        If Not namedRange Is Nothing Then
                            On Error Resume Next
                            Set targetRange = namedRange.RefersToRange
                            If Not targetRange Is Nothing Then
                                targetRange.Value = "In-Place"
                            End If
                            On Error GoTo 0
                        End If
                        
                        ' Instead of ws.Copy, manually copy values and formatting
                        Set newWS = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.count))
                        safeSheetName = "CF_" & Replace(Replace(ws.Name, ":", "_"), "\", "_")
                        On Error Resume Next
                        newWS.Name = Left(safeSheetName, 31) ' Max sheet name length = 31
                        On Error GoTo 0

                        Set usedRange = ws.usedRange
                        With usedRange
                            newWS.Range("A1").Resize(.Rows.count, .Columns.count).Value = .Value
                            .Copy
                            newWS.Range("A1").PasteSpecial Paste:=xlPasteFormats
                        End With

                        For col = 1 To usedRange.Columns.count
                            newWS.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth
                        Next col
                    End If
                Next ws

                wbSource.Close SaveChanges:=False
NextFile:
            End If
        Next file
    Next subfolderObj

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Cashflow sheets imported successfully!", vbInformation
End Sub

Sub ImportCashflowSheetsFromUWFiles1()
    Dim mainFolder As String
    Dim fso As Object, folder As Object, subfolderObj As Object
    Dim file As Object, wbSource As Workbook
    Dim ws As Worksheet, newWS As Worksheet
    Dim currentWB As Workbook
    Dim sheetCopied As Boolean
    Dim safeSheetName As String
    Dim namedRange As Name
    Dim targetRange As Range
    Dim usedRange As Range
    Dim col As Long
    
    ' Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Main Folder"
        If .Show <> -1 Then Exit Sub
        mainFolder = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(mainFolder)
    Set currentWB = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through each subfolder
    For Each subfolderObj In folder.SubFolders
        For Each file In subfolderObj.Files
            If file.Name Like "UW*.xls*" Then
                On Error Resume Next
                Set wbSource = Workbooks.Open(file.path, ReadOnly:=True)
                If Err.Number <> 0 Then
                    Debug.Print "Failed to open: " & file.path
                    Err.Clear
                    GoTo NextFile
                End If
                On Error GoTo 0
                
                For Each ws In wbSource.Worksheets
                    If LCase(ws.Name) Like "cash flow" Then
                                            
                        On Error Resume Next
                        Set namedRange = wbSource.Names("sizingcf")
                        On Error GoTo 0
                        
                        If Not namedRange Is Nothing Then
                            On Error Resume Next
                            Set targetRange = namedRange.RefersToRange
                            If Not targetRange Is Nothing Then
                                targetRange.Value = "In-Place"
                            End If
                            On Error GoTo 0
                        End If
                        
                        ws.Copy After:=currentWB.Sheets(currentWB.Sheets.count)
                        Set newWS = currentWB.Sheets(currentWB.Sheets.count)
                        safeSheetName = "CF_" & Replace(Replace(ws.Name, ":", "_"), "\", "_")
                        On Error Resume Next
                        newWS.Name = Left(safeSheetName, 31) ' Max sheet name length = 31
                        On Error GoTo 0
                    End If
                Next ws

                wbSource.Close SaveChanges:=False
NextFile:
            End If
        Next file
    Next subfolderObj

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Cashflow sheets imported successfully!", vbInformation
End Sub
Sub ImportCashflowSheetsFromUWFiles_ValuesOnly()
    Dim mainFolder As String
    Dim fso As Object, folder As Object, subfolderObj As Object
    Dim file As Object, wbSource As Workbook
    Dim ws As Worksheet, newWS As Worksheet
    Dim currentWB As Workbook
    Dim safeSheetName As String
    Dim namedRange As Name
    Dim targetRange As Range
    Dim usedRange As Range
    Dim col As Long

    ' Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Main Folder"
        If .Show <> -1 Then Exit Sub
        mainFolder = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(mainFolder)
    Set currentWB = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through each subfolder
    For Each subfolderObj In folder.SubFolders
        For Each file In subfolderObj.Files
            If file.Name Like "UW*.xls*" Then
                On Error Resume Next
                Set wbSource = Workbooks.Open(file.path, ReadOnly:=True)
                If Err.Number <> 0 Then
                    Debug.Print "Failed to open: " & file.path
                    Err.Clear
                    GoTo NextFile
                End If
                On Error GoTo 0

                ' Loop through sheets to find "Cash Flow"
                For Each ws In wbSource.Worksheets
                    If LCase(ws.Name) Like "cash flow" Then
                        ' Try to update named range "sizingcf"
                        On Error Resume Next
                        Set namedRange = wbSource.Names("sizingcf")
                        On Error GoTo 0
                        
                        If Not namedRange Is Nothing Then
                            On Error Resume Next
                            Set targetRange = namedRange.RefersToRange
                            If Not targetRange Is Nothing Then
                                targetRange.Value = "In-Place"
                            End If
                            On Error GoTo 0
                        End If

                        ' Create a new sheet in current workbook
                        Set newWS = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.count))
                        safeSheetName = "CF_" & Replace(Replace(ws.Name, ":", "_"), "\", "_")
                        On Error Resume Next
                        newWS.Name = Left(safeSheetName, 31)
                        On Error GoTo 0

                        ' Copy values and formatting from used range
                        Set usedRange = ws.usedRange
                        With usedRange
                            newWS.Range("A1").Resize(.Rows.count, .Columns.count).Value = .Value
                            .Copy
                            newWS.Range("A1").PasteSpecial Paste:=xlPasteFormats
                        End With

                        ' Copy column widths
                        For col = 1 To usedRange.Columns.count
                            newWS.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth
                        Next col

                        Exit For ' Stop after finding the Cash Flow sheet
                    End If
                Next ws

                wbSource.Close SaveChanges:=False
NextFile:
            End If
        Next file
    Next subfolderObj

    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Cashflow sheets copied with 'In-Place' update and values only.", vbInformation
End Sub
Sub ImportCashflowSheetsFromUWFiles_ValuesOnly2()
    Dim mainFolder As String
    Dim fso As Object, folder As Object, subfolderObj As Object
    Dim file As Object, wbSource As Workbook
    Dim ws As Worksheet, newWS As Worksheet
    Dim currentWB As Workbook
    Dim safeSheetName As String
    Dim namedRange As Name
    Dim targetRange As Range
    Dim usedRange As Range
    Dim col As Long
    Dim r As Long, c As Long
    Dim srcRange As Range

    ' Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the Main Folder"
        If .Show <> -1 Then Exit Sub
        mainFolder = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(mainFolder)
    Set currentWB = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through each subfolder
    For Each subfolderObj In folder.SubFolders
        For Each file In subfolderObj.Files
            If file.Name Like "UW*.xls*" Then
                On Error Resume Next
                Set wbSource = Workbooks.Open(file.path, ReadOnly:=True)
                If Err.Number <> 0 Then
                    Debug.Print "Failed to open: " & file.path
                    Err.Clear
                    GoTo NextFile
                End If
                On Error GoTo 0

                ' Loop through sheets to find "Cash Flow"
                For Each ws In wbSource.Worksheets
                    If LCase(ws.Name) Like "cash flow" Then
                        ' Try to update named range "sizingcf"
                        On Error Resume Next
                        Set namedRange = wbSource.Names("sizingcf")
                        On Error GoTo 0
                        
                        If Not namedRange Is Nothing Then
                            On Error Resume Next
                            Set targetRange = namedRange.RefersToRange
                            If Not targetRange Is Nothing Then
                                targetRange.Value = "In-Place"
                            End If
                            On Error GoTo 0
                        End If

                        ' Create a new sheet in current workbook
                        Set newWS = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.count))
                        safeSheetName = "CF_" & Replace(Replace(ws.Name, ":", "_"), "\", "_")
                        On Error Resume Next
                        newWS.Name = Left(safeSheetName, 31)
                        On Error GoTo 0

                        ' Copy values cell by cell to avoid overflow
                        Set usedRange = ws.usedRange
                        Set srcRange = usedRange

                        For r = 1 To srcRange.Rows.count
                            For c = 1 To srcRange.Columns.count
                                newWS.Cells(r, c).Value = srcRange.Cells(r, c).Value
                            Next c
                        Next r

                        ' Copy formatting
                        srcRange.Copy
                        newWS.Range("A1").PasteSpecial Paste:=xlPasteFormats
                        Application.CutCopyMode = False

                        ' Copy column widths
                        For col = 1 To srcRange.Columns.count
                            newWS.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth
                        Next col

                        Exit For ' Stop after finding the Cash Flow sheet
                    End If
                Next ws

                wbSource.Close SaveChanges:=False
NextFile:
            End If
        Next file
    Next subfolderObj

    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Cashflow sheets copied with 'In-Place' update and values only.", vbInformation
End Sub




Sub CopyCashFlowForEachDropdownValue()
    Dim filePath As String
    Dim wbSource As Workbook
    Dim wsCashFlow As Worksheet
    Dim validationList As String
    Dim listItems As Variant
    Dim i As Long
    Dim currentWB As Workbook
    Dim newSheet As Worksheet
    Dim safeSheetName As String
    Dim refSheetName As String
    Dim refRangeAddress As String
    Dim refSheet As Worksheet
    Dim refRange As Range

    ' Ask user to select the Excel file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Excel file with dropdown in D4"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Set currentWB = ThisWorkbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Open the selected workbook
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)

    ' Check if "Cash Flow" sheet exists
    On Error Resume Next
    Set wsCashFlow = wbSource.Sheets("Cash Flow")
    On Error GoTo 0
    If wsCashFlow Is Nothing Then
        MsgBox "'Cash Flow' sheet not found!", vbExclamation
        wbSource.Close SaveChanges:=False
        Exit Sub
    End If

    ' Get dropdown values from D4
    With wsCashFlow.Range("D4").Validation
        If .Type = xlValidateList Then
            validationList = .Formula1
        Else
            MsgBox "No dropdown list found in D4!", vbExclamation
            wbSource.Close SaveChanges:=False
            Exit Sub
        End If
    End With

    ' Handle reference to another sheet like =Summary!$F$4:$F$20
    If Left(validationList, 1) = "=" Then
        validationList = Mid(validationList, 2) ' remove '='

        If InStr(validationList, "!") > 0 Then
            ' Split into sheet and range
            refSheetName = Split(validationList, "!")(0)
            refRangeAddress = Split(validationList, "!")(1)
            
            On Error Resume Next
            Set refSheet = wbSource.Sheets(refSheetName)
            On Error GoTo 0

            If refSheet Is Nothing Then
                MsgBox "Could not find the sheet: " & refSheetName, vbExclamation
                wbSource.Close SaveChanges:=False
                Exit Sub
            End If

            On Error Resume Next
            Set refRange = refSheet.Range(refRangeAddress)
            On Error GoTo 0

            If refRange Is Nothing Then
                MsgBox "Could not resolve the list range: " & refRangeAddress, vbExclamation
                wbSource.Close SaveChanges:=False
                Exit Sub
            End If

            listItems = Application.Transpose(refRange.Value)
        Else
            ' Named range
            On Error Resume Next
            Set refRange = wbSource.Range(validationList)
            On Error GoTo 0

            If refRange Is Nothing Then
                MsgBox "Could not resolve named range: " & validationList, vbExclamation
                wbSource.Close SaveChanges:=False
                Exit Sub
            End If

            listItems = Application.Transpose(refRange.Value)
        End If
    Else
        ' Comma-separated list
        listItems = Split(validationList, ",")
    End If

    ' Loop through each dropdown value
    For i = LBound(listItems) To UBound(listItems)
        If Trim(listItems(i)) <> "" Then
            wsCashFlow.Range("D4").Value = listItems(i)
            wsCashFlow.Calculate
            
            wsCashFlow.Copy After:=currentWB.Sheets(currentWB.Sheets.count)
            Set newSheet = currentWB.Sheets(currentWB.Sheets.count)
            
            safeSheetName = "Portfolio_" & Replace(Replace(CStr(listItems(i)), ":", "_"), "\", "_")
            On Error Resume Next
            newSheet.Name = Left(safeSheetName, 31)
            On Error GoTo 0
        End If
    Next i

    wbSource.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "All 'Cash Flow' sheets copied with Portfolio prefix!", vbInformation
End Sub

Sub CopyCashFlowValuesOnlyForEachDropdownValue()
    Dim filePath As String
    Dim wbSource As Workbook
    Dim wsCashFlow As Worksheet
    Dim validationList As String
    Dim listItems As Variant
    Dim i As Long
    Dim currentWB As Workbook
    Dim newSheet As Worksheet
    Dim safeSheetName As String
    Dim refSheetName As String
    Dim refRangeAddress As String
    Dim refSheet As Worksheet
    Dim refRange As Range
    Dim usedRange As Range
    Dim portfoliorange As String
    'Set portfoliorange = Nothing

    ' Ask user to select the Excel file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Excel file with dropdown in D4"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Set currentWB = ThisWorkbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Open the selected workbook
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)

    ' Check if "Cash Flow" sheet exists
    On Error Resume Next
    Set wsCashFlow = wbSource.Sheets("Cash Flow")
    On Error GoTo 0
    If wsCashFlow Is Nothing Then
        MsgBox "'Cash Flow' sheet not found!", vbExclamation
        wbSource.Close SaveChanges:=False
        Exit Sub
    End If
    
'    On Error Resume Next
 '   portfoliorange = GetRange(wsCashFlow, "Portfolio", "A1", 7, 7)
  '  On Error GoTo 0
   '
    'If portfoliorange Is Nothing Then
     '   MsgBox "Could not find the cell with text 'Portfolio'.", vbCritical
      '  wbSource.Close SaveChanges:=False
       ' Exit Sub
'    Else
 '       MsgBox "Found 'Portfolio' at: " & portfoliorange.address, vbInformation
  '  End If

    ' Get dropdown values from D4
    Dim cellAddress As String
    On Error Resume Next
    cellAddress = Application.InputBox( _
        Prompt:="Enter the cell address that contains the dropdown (e.g., C3):", _
        Title:="Enter Cell Address", Type:=2) ' Type 2 = Text input
    On Error GoTo 0
    
    If cellAddress = "" Then
        MsgBox "No cell address entered. Operation cancelled.", vbExclamation
        wbSource.Close SaveChanges:=False
        Exit Sub
    End If


    portfoliorange = cellAddress
    'If Not portfoliorange Is Nothing Then
        With wsCashFlow.Range(portfoliorange).Validation
            If .Type = xlValidateList Then
                validationList = .Formula1
            Else
                MsgBox "No dropdown list found in D4!", vbExclamation
                wbSource.Close SaveChanges:=False
                Exit Sub
            End If
        End With
    'End If
    

    ' Handle reference to another sheet like =Summary!$F$4:$F$20
    If Left(validationList, 1) = "=" Then
        validationList = Mid(validationList, 2)

        If InStr(validationList, "!") > 0 Then
            refSheetName = Split(validationList, "!")(0)
            refRangeAddress = Split(validationList, "!")(1)
            Set refSheet = wbSource.Sheets(refSheetName)
            Set refRange = refSheet.Range(refRangeAddress)
            listItems = Application.Transpose(refRange.Value)
        Else
            Set refRange = wbSource.Range(validationList)
            listItems = Application.Transpose(refRange.Value)
        End If
    Else
        listItems = Split(validationList, ",")
    End If

    ' Loop through dropdown values
    For i = LBound(listItems) To UBound(listItems)
        If Trim(listItems(i)) <> "" Then
            wsCashFlow.Range(portfoliorange).Value = listItems(i)
            'portfoliorange.Value = listItems(i)
            wsCashFlow.Calculate

            ' Create new sheet in destination workbook
            Set newSheet = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.count))
            safeSheetName = "Portfolio_" & Replace(Replace(CStr(listItems(i)), ":", "_"), "\", "_")
            On Error Resume Next
            newSheet.Name = Left(safeSheetName, 31)
            On Error GoTo 0

            ' Copy values and formatting
            Set usedRange = wsCashFlow.usedRange
            With usedRange
                newSheet.Range("A1").Resize(.Rows.count, .Columns.count).Value = .Value
                .Copy
                newSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
            End With

            ' Copy column widths
            Dim col As Long
            For col = 1 To usedRange.Columns.count
                newSheet.Columns(col).ColumnWidth = wsCashFlow.Columns(col).ColumnWidth
            Next col
        End If
    Next i

    wbSource.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Cash Flow values copied with formatting only (no formulas).", vbInformation
End Sub



Sub CreateAssetSheet()
    Dim ws As Worksheet
    Dim assetWS As Worksheet
    Dim rowNum As Long
    Dim sheetIndex As Integer
    Dim loanID As String
    Dim assetID As String, assetName As String
    Dim sqftUnit As Variant, sqFt As Variant, units As Variant
    Dim yearBuilt As Variant, yearRenovated As Variant, appraisedValue As Variant
    Dim noi As Variant, capRate As Variant, locationType As Variant

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Delete existing Asset sheet
    On Error Resume Next: Worksheets("Asset").Delete: On Error GoTo 0
    Set assetWS = ThisWorkbook.Sheets.Add
    assetWS.Name = "Asset"

    ' Add headers
    With assetWS
        .Range("A1:S1").Value = Array("Loan ID", "Asset ID", "Asset Loan Allocation", "Asset Name", "Asset Address", _
        "Square Footage/Unit", "Square Footage", "Units", "Main Type of Use", "Year Built", "Year Renovate", _
        "Appraised Value", "Appraised Value Date", "Net Operating Income", "Location Type", "Class", _
        "Type of Use Detailed Description", "Cap Rate", "Portfolio")
        .Rows(1).Font.Bold = True
    End With

    rowNum = 2

    ' Loop through all sheets starting from the 3rd one
    For sheetIndex = 3 To ThisWorkbook.Sheets.count
        Set ws = ThisWorkbook.Sheets(sheetIndex)

        loanID = ws.Name ' or extract from cell if needed
        assetID = ""     ' Define logic to fetch Asset ID if needed
        assetName = FindAdjacentValue(ws, "Property Name", "right", ws.Range("A1:M200"), 5, 5)  ' Define logic to fetch Asset Name if needed
        If assetName = "Not Found" Or assetName = "No Value Found" Or assetName = "Invalid Direction" Then
            assetName = FindAdjacentValueX(ws, "Select Property", "right", ws.Range("A1:M100"), 5, 5)
        End If

        yearBuilt = FindAdjacentValue(ws, "Year Built", "right", ws.Range("A1:M200"), 5, 5)
        yearRenovated = FindAdjacentValue(ws, "Year Renovated", "right", ws.Range("A1:M200"), 5, 5)
        appraisedValue = FindAdjacentValue(ws, "Appraised Value", "right", ws.Range("A1:M200"), 5, 5)
        noi = FindAdjacentValue(ws, "NET OPERATING INCOME", "right", ws.Range("A1:M200"), 5, 5)
        capRate = FindAdjacentValue(ws, "Cap Rate", "right", ws.Range("H1:AC100"), 5, 5)
        
        
        sqftUnit = FindAdjacentValueX(ws, "Total SF / Adjusted SF  (A)", "right", ws.Range("A1:M100"), 5, 5)
        If sqftUnit = "Not Found" Or sqftUnit = "No Value Found" Or sqftUnit = "Invalid Direction" Then
            sqftUnit = FindAdjacentValueX(ws, "Tot. Units / Adj. Tot. Units (A)", "right", ws.Range("A1:M100"), 5, 5)
        End If
        If IsError(sqftUnit) Then sqftUnit = "ds"
        If sqftUnit = "Not Found" Or sqftUnit = "No Value Found" Or sqftUnit = "Invalid Direction" Then
            sqftUnit = FindAdjacentValueX(ws, "Units", "right", ws.Range("A1:M100"), 5, 5)
        End If
        If IsError(sqftUnit) Then sqftUnit = "ds"
        If sqftUnit = "Not Found" Or sqftUnit = "No Value Found" Or sqftUnit = "Invalid Direction" Then
            sqftUnit = FindAdjacentValueX(ws, "No. Units", "right", ws.Range("A1:M100"), 5, 5)
        End If

        
        
        
        
        sqFt = "" ' Placeholder – update logic if needed
        
        sqFt = FindAdjacentValueX(ws, "Total SF / Adjusted SF  (A)", "right", ws.Range("A1:M100"), 5, 5)
        If sqFt = "Not Found" Or sqFt = "No Value Found" Or sqFt = "Invalid Direction" Then
            sqFt = FindAdjacentValueX(ws, "Tot. Units / Adj. Tot. Units (A)", "right", ws.Range("A1:M100"), 5, 5)
        End If
        If IsError(sqFt) Then sqFt = "ds"
        If sqFt = "Not Found" Or sqFt = "No Value Found" Or sqFt = "Invalid Direction" Then
            sqFt = FindAdjacentValueX(ws, "Units", "right", ws.Range("A1:M100"), 5, 5)
        End If
        If IsError(sqFt) Then sqFt = "ds"
        If sqFt = "Not Found" Or sqFt = "No Value Found" Or sqFt = "Invalid Direction" Then
            sqFt = FindAdjacentValueX(ws, "No. Units", "right", ws.Range("A1:M100"), 5, 5)
        End If


        units = "" ' Placeholder – update logic if needed
        locationType = ""

        With assetWS
            .Cells(rowNum, 1).Value = loanID
            .Cells(rowNum, 2).Value = assetID
            .Cells(rowNum, 3).Formula = "=IFERROR(L" & rowNum & "/SUMIF($A:$A,$A" & rowNum & ",$L:$L),0)"
            .Cells(rowNum, 4).Value = assetName
            .Cells(rowNum, 5).Value = "" ' Address placeholder
            .Cells(rowNum, 6).Value = sqftUnit
            .Cells(rowNum, 7).Value = sqFt
            .Cells(rowNum, 8).Value = units
            .Cells(rowNum, 9).Value = "" ' Type of use
            .Cells(rowNum, 10).Value = yearBuilt
            .Cells(rowNum, 11).Value = yearRenovated
            .Cells(rowNum, 12).Value = appraisedValue
            .Cells(rowNum, 13).Value = "" ' Appraised value date
            .Cells(rowNum, 14).Value = noi
            .Cells(rowNum, 15).Value = locationType
            .Cells(rowNum, 16).Value = "" ' Class
            .Cells(rowNum, 17).Value = "" ' Type of Use Description
            If IsNumeric(capRate) And capRate <> "" Then
                .Cells(rowNum, 18).Value = Format(capRate, "0.00%")
            Else
                .Cells(rowNum, 18).Value = capRate
            End If
            .Cells(rowNum, 19).Value = "" ' Portfolio
        End With

        rowNum = rowNum + 1
    Next sheetIndex

    assetWS.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Asset sheet created from all sheets starting from the 3rd one!", vbInformation
End Sub
Function GetRange(ws As Worksheet, searchText As String, startAddress As String, maxRight As Integer, maxDown As Integer) As String
    Dim startCell As Range
    Dim r As Long, c As Long
    Dim currentCell As Range
    
    Set startCell = ws.Range(startAddress)

    For r = 0 To maxDown
        For c = 0 To maxRight
            Set currentCell = startCell.Offset(r, c)
            If Not IsError(currentCell.Value) Then
                If Trim(CStr(currentCell.Value)) = Trim(searchText) Then
                    FindCellAddressInRange = currentCell.address
                    Exit Function
                End If
            End If
        Next c
    Next r

    FindCellAddressInRange = "" ' Return empty string if not found
End Function




Function FindAdjacentValue(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim foundCell As Range
    Dim checkCell As Range
    Dim r As Long, c As Long
    Dim i As Integer
    
    ' Search for the searchText in the range
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

    On Error GoTo CleanExit

    If direction = "right" Then
        For i = 1 To maxRight
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
        Next i
        FindAdjacentValue = "No Value Found"

    ElseIf direction = "down" Then
        For i = 1 To maxDown
            Set checkCell = ws.Cells(r + i, c)
            If Trim(checkCell.Value) <> "" Then
                FindAdjacentValue = checkCell.Value
                Exit Function
            End If
        Next i
        FindAdjacentValue = "No Value Found"
        
    Else
        FindAdjacentValue = "Invalid Direction"
    End If

CleanExit:
    On Error GoTo 0
End Function

Sub TestFind()
    Dim result As Variant
    Dim Sheet1 As Worksheet
    Set Sheet1 = ThisWorkbook.Sheets("Aurora Marketplace")
    
    result = FindAdjacentValueX(Sheet1, "Total SF / Adjusted SF  (A)", "right", Sheet1.Range("A1:M200"), 5, 5)
    MsgBox "Result: " & result
End Sub

Function FindAdjacentValueX(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim foundCell As Range
    Dim i As Integer
    Dim checkCell As Range

    ' Clean search text
    searchText = CleanString(searchText)

    ' Step 1: Find the cell with the exact (cleaned) value
    For Each cell In searchRange
        If CleanString(CStr(cell.Value)) = searchText Then
            Set foundCell = cell
            Exit For
        End If
    Next cell

    If foundCell Is Nothing Then
        FindAdjacentValueX = "Not Found"
        Exit Function
    End If

    On Error GoTo CleanExit
    ' Step 2: Get value from the next cell in the specified direction
    If LCase(direction) = "right" Then
        For i = 1 To maxRight
            Set checkCell = ws.Cells(foundCell.Row, foundCell.Column + i)
            If checkCell.MergeCells Then Set checkCell = checkCell.MergeArea.Cells(1, 2)
            If Trim(CStr(checkCell.Value)) <> "" Then
                FindAdjacentValueX = checkCell.Value
                Exit Function
            End If
        Next i

    ElseIf LCase(direction) = "down" Then
        For i = 1 To maxDown
            Set checkCell = ws.Cells(foundCell.Row + i, foundCell.Column)
            If checkCell.MergeCells Then Set checkCell = checkCell.MergeArea.Cells(1, 1)
            If Trim(CStr(checkCell.Value)) <> "" Then
                FindAdjacentValueX = checkCell.Value
                Exit Function
            End If
        Next i
    Else
        FindAdjacentValueX = "Invalid Direction"
        Exit Function
    End If

    FindAdjacentValueX = "No Value Found"
    
CleanExit:
    On Error GoTo 0
End Function
Function FindAdjacentValueXY(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim foundCell As Range
    Dim i As Integer
    Dim checkCell As Range

    ' Clean search text
    searchText = CleanString(searchText)

    ' Step 1: Find the cell with the exact (cleaned) value
    For Each cell In searchRange
        If Not IsError(cell.Value) Then
            If CleanString(CStr(cell.Value)) = searchText Then
                Set foundCell = cell
                Exit For
            End If
        End If
    Next cell

    If foundCell Is Nothing Then
        FindAdjacentValueX = "Not Found"
        Exit Function
    End If

    On Error GoTo CleanExit

    ' Step 2: Get value from the next cell in the specified direction
    If LCase(direction) = "right" Then
        For i = 1 To maxRight
            Set checkCell = ws.Cells(foundCell.Row, foundCell.Column + i)
            If checkCell.MergeCells Then Set checkCell = checkCell.MergeArea.Cells(1, 1)
            If Not IsError(checkCell.Value) Then
                If Trim(CStr(checkCell.Value)) <> "" Then
                    FindAdjacentValueX = checkCell.Value
                    Exit Function
                End If
            End If
        Next i

    ElseIf LCase(direction) = "down" Then
        For i = 1 To maxDown
            Set checkCell = ws.Cells(foundCell.Row + i, foundCell.Column)
            If checkCell.MergeCells Then Set checkCell = checkCell.MergeArea.Cells(1, 1)
            If Not IsError(checkCell.Value) Then
                If Trim(CStr(checkCell.Value)) <> "" Then
                    FindAdjacentValueX = checkCell.Value
                    Exit Function
                End If
            End If
        Next i
    Else
        FindAdjacentValueX = "Invalid Direction"
        GoTo CleanExit
    End If

    FindAdjacentValueX = "No Value Found"

CleanExit:
    On Error GoTo 0
End Function




Function CleanString(str As String) As String
    ' Remove non-breaking spaces, tabs, newlines, extra spaces
    str = Replace(str, Chr(160), "")
    str = Replace(str, vbTab, "")
    str = Replace(str, vbCr, "")
    str = Replace(str, vbLf, "")
    CleanString = Trim(str)
End Function



Function FindAdjacentValue1(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
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


