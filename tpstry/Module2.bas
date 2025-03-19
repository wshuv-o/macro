Attribute VB_Name = "Module2"
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
    Dim fso As Object
    Dim data As Variant
    Dim i As Long
    
    ' Set the destination sheet for output
    Set destSheet = ThisWorkbook.Sheets("Tracker")
    incrementRow = 2

    ' Select a folder containing subfolders/files.xlsm
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Subfolders"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    
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
                        Do While Not IsEmpty(loanSummarySheet.Cells(loanSummaryRow, 6).value) And _
                            Not loanSummarySheet.Cells(loanSummaryRow, 6).value Like "*Total*"

                            ' Populate columns A, B, C, D, E, F and G

                            destSheet.Cells(incrementRow, 1).value = subFolderPart1
                            destSheet.Cells(incrementRow, 2).value = subFolderPart1 & "-" & (increment)
                            destSheet.Cells(incrementRow, 3).value = subFolderPart2
                            destSheet.Cells(incrementRow, 4).value = loanSummarySheet.Cells(loanSummaryRow, 6).value ' D - Asset name from Cashflow
                            destSheet.Cells(incrementRow, 5).value = _
                                loanSummarySheet.Cells(loanSummaryRow, 20).value & ", " & _
                                loanSummarySheet.Cells(loanSummaryRow, 22).value & ", " & _
                                loanSummarySheet.Cells(loanSummaryRow, 23).value & " " & _
                                loanSummarySheet.Cells(loanSummaryRow, 24).value                                     ' E - Address concat(T, V, W, X)
                            destSheet.Cells(incrementRow, 6).value = loanSummarySheet.Cells(loanSummaryRow, 9).value ' F - Loan Summary sheet's column 9 (I)
                            destSheet.Cells(incrementRow, 7).value = subFolder.Name                                  ' G - Loan Name from Folder
                            
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

    ' Re-enable screen updating, calculation, events, and status bar
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    MsgBox "Data extraction complete!", vbInformation
    Exit Sub

FileError:
    MsgBox "Error opening file: " & file.Path, vbCritical
    Resume Next

End Sub



Sub PullDataFromFolder()
    '-----------------Tracker pull Vars-------------------
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
    Dim subFolderPart1 As Variant
    Dim subFolderPart2 As Variant
    Dim spacePos As Integer
    Dim fso As Object
    '-----------------Cashflow pull Vars-------------------
    Dim netCashFlowCell As Range
    Dim copiedRange As Range
    Dim lastRow As Long
    Dim dataArray As Variant
    Dim sheet As Worksheet
    Dim newSheet As Worksheet
    Dim newSheetName As String
    Dim invalidChars As String
    Dim i As Integer
    Dim sheetCounter As Integer
    '-------------------Asset Sheet---------------
    Dim assetSheet As Worksheet
    Dim aggregateCashFlow As Worksheet
    Dim incrementRowAsset As Integer
    Dim loanSheet As Worksheet
    Dim incrementRowLoan As Integer
    
    '--------------------
    Dim lastRowTrackerSheet As Integer
    Dim foundCell As Range
    
    ' Set destination sheet for output
    Set destSheet = ThisWorkbook.Sheets("Tracker")
    Set assetSheet = ThisWorkbook.Sheets("Asset")
    Set loanSheet = ThisWorkbook.Sheets("Loan")
    incrementRow = 2
    incrementRowAsset = 6
    incrementRowLoan = 6

    ' Select folder containing subfolders/files.xlsm
    With Application.fileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Subfolders"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
    End With

    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Disable screen updating, calculations, and events to improve performance
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    
    ' Loop through all subfolders in the selected folder
    For Each subFolder In fso.GetFolder(folderPath).SubFolders
        subFolderName = subFolder.Name
        spacePos = InStr(subFolderName, " ") ' Find the position of the first space

        If spacePos > 0 Then
            ' Split the subfolder name into two parts (before and after the first space)
            subFolderPart1 = Left(subFolderName, spacePos - 1)
            subFolderPart2 = Mid(subFolderName, spacePos + 1)
            
            ' Loop through each file in the subfolder
            For Each file In subFolder.Files
                fileName = file.Name

                ' Check if the file starts with 'UW' and is an Excel file
                If fileName Like "UW*" And _
                   (Right(fileName, 4) = ".xls" Or Right(fileName, 5) = ".xlsx" Or Right(fileName, 5) = ".xlsm") Then

                    ' Open the workbook for Tracker data
                    Set wb = Workbooks.Open(file.Path, ReadOnly:=True)
                    Set loanSummarySheet = wb.Sheets("Loan Analysis")
                    
                    
                    ' Process Tracker Details if the sheet exists
                    If Not loanSummarySheet Is Nothing Then
                        loanSummaryRow = 66
                        increment = 1 ' Start from row 2 to avoid overwriting header
                        
                        ' Loop through rows starting from 66
                        Do While Not IsEmpty(loanSummarySheet.Cells(loanSummaryRow, 6).value) And _
                            Not loanSummarySheet.Cells(loanSummaryRow, 6).value Like "*Total*"
    
                            ' Populate Tracker data columns A to G
                            destSheet.Cells(incrementRow, 1).value = subFolderPart1
                            destSheet.Cells(incrementRow, 2).value = subFolderPart1 & "-" & increment
                            destSheet.Cells(incrementRow, 3).value = subFolderPart2
                            destSheet.Cells(incrementRow, 4).value = loanSummarySheet.Cells(loanSummaryRow, 6).value ' Asset name
                            destSheet.Cells(incrementRow, 5).value = _
                                loanSummarySheet.Cells(loanSummaryRow, 20).value & ", " & _
                                loanSummarySheet.Cells(loanSummaryRow, 22).value & ", " & _
                                loanSummarySheet.Cells(loanSummaryRow, 23).value & " " & _
                                loanSummarySheet.Cells(loanSummaryRow, 24).value ' Address
                            destSheet.Cells(incrementRow, 6).value = loanSummarySheet.Cells(loanSummaryRow, 9).value ' Loan Summary
                            destSheet.Cells(incrementRow, 7).value = subFolder.Name ' Loan Name from Folder
                            destSheet.Cells(incrementRow, 9).value = "=OFFSET(Mapping!$C$4, MATCH(F" & incrementRow & ", Mapping!$B$5:$B$60, 0), 0)"
                            
                            
                            
                            
                            
                            ' Populate Asset data
                            assetSheet.Cells(incrementRowAsset, 1).value = subFolderPart1
                            assetSheet.Cells(incrementRowAsset, 2).value = subFolderPart1 & "-" & increment
                            assetSheet.Cells(incrementRowAsset, 3).value = "=IFERROR(K" & incrementRowAsset & "/SUMIF($A:$A,$A" & incrementRowAsset & ",$K:$K),0)"
                            assetSheet.Cells(incrementRowAsset, 4).value = loanSummarySheet.Cells(loanSummaryRow, 6).value ' Asset name
                            assetSheet.Cells(incrementRowAsset, 5).value = destSheet.Cells(incrementRow, 5).value
                            assetSheet.Cells(incrementRowAsset, 6).value = loanSummarySheet.Cells(loanSummaryRow, 10).value
                            assetSheet.Cells(incrementRowAsset, 7).value = loanSummarySheet.Cells(loanSummaryRow, 10).value
                            assetSheet.Cells(incrementRowAsset, 8).value = loanSummarySheet.Cells(loanSummaryRow, 9).value
                            assetSheet.Cells(incrementRowAsset, 9).value = loanSummarySheet.Cells(loanSummaryRow, 25).value
                            assetSheet.Cells(incrementRowAsset, 10).value = loanSummarySheet.Cells(loanSummaryRow, 26).value
                            assetSheet.Cells(incrementRowAsset, 11).value = loanSummarySheet.Cells(loanSummaryRow, 16).value    'Appraisal Value
                            'assetSheet.Cells(incrementRowAsset, 12).value = ""                                                  'Appraisal value date
                            assetSheet.Cells(incrementRowAsset, 13).value = loanSummarySheet.Cells(loanSummaryRow, 11).value    'NOI
                            'assetSheet.Cells(incrementRowAsset, 14).value = ""                                                  'Net Operating Income at Origination
                            'assetSheet.Cells(incrementRowAsset, 15).value = ""                                                  'Location Type
                            'assetSheet.Cells(incrementRowAsset, 16).value = ""                                                  'Class
                            assetSheet.Cells(incrementRowAsset, 17).value = loanSummarySheet.Cells(loanSummaryRow, 9).value     'Type of Use Detailed Description
                            assetSheet.Cells(incrementRowAsset, 18).value = loanSummarySheet.Cells(loanSummaryRow, 13).value    'Cap Rate
                            'assetSheet.Cells(incrementRowAsset, 19).value = ""                                                  'Portfolio
                            
                            
                            
                            
                            
                            
                            
                            ' Populate Loan data
                            lastRowTrackerSheet = ThisWorkbook.Sheets("Tracker").Cells(Rows.Count, 1).End(xlUp).Row
                            If increment = 1 Then
                                ' If the Loan ID in the current row is different, insert a new row and fill the data
                                loanSheet.Cells(incrementRowLoan, 2).value = "=TEXTJOIN("", "", TRUE, FILTER(Tracker!B:B, Tracker!A:A=A" & incrementRowLoan & "))" ' Associated Asset(s) ID
                                loanSheet.Cells(incrementRowLoan, 3).value = loanSummarySheet.Range("LS_NoteDate").value                                ' Note Date
                                loanSheet.Cells(incrementRowLoan, 4).value = loanSummarySheet.Range("LS_LoanAmount").value                      ' Original Loan Amount
                                loanSheet.Cells(incrementRowLoan, 5).value = loanSheet.Cells(incrementRowLoan, 4).value                         ' Current Loan Amount
                                loanSheet.Cells(incrementRowLoan, 6).value = "=EOMONTH(C" & incrementRowLoan & ",AG" & incrementRowLoan & ")+1" ' Maturity Date
                                
                                Set foundCell = loanSummarySheet.Columns(6).Find(What:="Debt Service", LookAt:=xlWhole)                         'Annual Debt Service Loan Analysis!K123
                                If Not foundCell Is Nothing Then
                                    loanSheet.Cells(incrementRowLoan, 7).value = loanSummarySheet.Cells(foundCell.Row, 11).value
                                Else
                                    loanSheet.Cells(incrementRowLoan, 7).value = "Not Found"
                                End If


                                loanSheet.Cells(incrementRowLoan, 8).value = "=G" & incrementRowLoan & "/12"                                    ' Monthly Debt Service
                                loanSheet.Cells(incrementRowLoan, 9).value = loanSummarySheet.Cells(12, 23).value                               ' Current Debt Yield* Loan Analysis!W12
                                loanSheet.Cells(incrementRowLoan, 10).value = loanSummarySheet.Cells(12, 20).value                              ' Current LTV* Loan Analysis!T12
                                loanSheet.Cells(incrementRowLoan, 11).value = "" ' Current DSCR
                                loanSheet.Cells(incrementRowLoan, 12).value = "" ' Current Loan KPI as of Date
                                loanSheet.Cells(incrementRowLoan, 13).value = "=EOMONTH(C" & incrementRowLoan & ",N" & incrementRowLoan & ")+1" ' Interest Only End Date
                                loanSheet.Cells(incrementRowLoan, 14).value = loanSummarySheet.Range("IOPeriods").value                         ' Interest Only Period
                                loanSheet.Cells(incrementRowLoan, 15).value = loanSummarySheet.Range("LoanAnalysis_Rate").value                 ' Interest Rate
                                loanSheet.Cells(incrementRowLoan, 16).value = loanSummarySheet.Range("IndexValue").value                        ' Interest Rate Index
                                loanSheet.Cells(incrementRowLoan, 17).value = loanSummarySheet.Range("LS_Spread").value                         ' Interest Rate Spread
                                loanSheet.Cells(incrementRowLoan, 18).value = loanSummarySheet.Range("LS_IndexType").value                      ' Interest Type
                                loanSheet.Cells(incrementRowLoan, 19).value = "" ' Commitment Amount
                                loanSheet.Cells(incrementRowLoan, 20).value = "" ' Contact Name
                                loanSheet.Cells(incrementRowLoan, 21).value = "" ' Contact Type
                                loanSheet.Cells(incrementRowLoan, 22).value = "" ' Debt Yield at Origination
                                loanSheet.Cells(incrementRowLoan, 23).value = "" ' DSCR at Origination
                                loanSheet.Cells(incrementRowLoan, 24).value = loanSummarySheet.Range("AmortTerm").value ' Amortization Term
                                loanSheet.Cells(incrementRowLoan, 25).value = "=TEXTJOIN("", "", TRUE, FILTER(Tracker!D:D, Tracker!A:A=A" & incrementRowLoan & "))" ' Borrower
                                loanSheet.Cells(incrementRowLoan, 26).value = loanSummarySheet.Range("FIPDate").value ' First Payment Date
                                loanSheet.Cells(incrementRowLoan, 27).value = "" ' Grace Period
                                loanSheet.Cells(incrementRowLoan, 28).value = "" ' Guarantor
                                loanSheet.Cells(incrementRowLoan, 29).value = "" ' Lender
                                loanSheet.Cells(incrementRowLoan, 30).value = "=INDEX(Tracker!C:C, MATCH(A" & incrementRowLoan & ", Tracker!A:A, 0))" ' Loan Name
                                loanSheet.Cells(incrementRowLoan, 31).value = "" ' Loan Product
                                loanSheet.Cells(incrementRowLoan, 32).value = loanSummarySheet.Range("LS_LoanPurpose").value                    ' Loan Purpose
                                loanSheet.Cells(incrementRowLoan, 33).value = loanSummarySheet.Range("LS_Term").value                           ' Loan Term
                                loanSheet.Cells(incrementRowLoan, 34).value = "" ' Loan Type
                                loanSheet.Cells(incrementRowLoan, 35).value = "" ' LTC
                                loanSheet.Cells(incrementRowLoan, 36).value = loanSummarySheet.Cells(12, 20).value                              ' LTV at Origination Loan Analysis!T12
                                loanSheet.Cells(incrementRowLoan, 37).value = "" ' Next Payment Date
                                loanSheet.Cells(incrementRowLoan, 38).value = "" ' Open Prepayment Date
                                loanSheet.Cells(incrementRowLoan, 39).value = "" ' Origination Status
                                loanSheet.Cells(incrementRowLoan, 40).value = "" ' Recourse
                                loanSheet.Cells(incrementRowLoan, 41).value = "" ' Recourse Description
                                loanSheet.Cells(incrementRowLoan, 42).value = "" ' Risk Rating
                                loanSheet.Cells(incrementRowLoan, 43).value = "" ' Servicer
                                loanSheet.Cells(incrementRowLoan, 44).value = "" ' Sponsor
                                loanSheet.Cells(incrementRowLoan, 45).value = "" ' Watchlist Indicator
                                
                                incrementRowLoan = incrementRowLoan + 1 ' Move to the next row
                            End If
                            
                            ' Increment row counters
                            increment = increment + 1
                            incrementRow = incrementRow + 1
                            incrementRowAsset = incrementRowAsset + 1
                            loanSummaryRow = loanSummaryRow + 1
                        Loop
                    End If
                    
                    

                    For Each sheet In wb.Sheets
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

                                ' Load data from copiedRange into an array for faster manipulation
                                dataArray = copiedRange.value
                                newSheetName = sheet.Range("H17").value

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

                                ''If the sheet already exists, append a counter until a unique name is found
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
                                Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                                newSheet.Name = newSheetName

                                ' Write the dataArray back to the new sheet in one go
                                newSheet.Range("A1").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).value = dataArray

                                ' Copy formats from the original range
                                copiedRange.Copy
                                newSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
                                Application.CutCopyMode = False
                            End If
                        End If
                    Next sheet
                    
                    ' Close the workbook without saving
                    wb.Close False

                End If
            Next file
        End If
    Next subFolder





    ' Re-enable screen updating, calculation, events, and status bar
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    
    MsgBox "Data extraction complete!", vbInformation
End Sub
