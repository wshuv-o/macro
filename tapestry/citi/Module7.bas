Attribute VB_Name = "Module7"
Sub CreateAssetSheet()
    Dim batchFolder As String
    Dim fso As Object, folder As Object, subFolder As Object
    Dim file As Object, wb As Workbook, ws As Worksheet
    Dim assetWS As Worksheet, rowNum As Long
    Dim folderName As String, filePath As String
    Dim loanID As String
    Dim assetCounter As Long, assetID As String
    Dim assetName As String
    Dim sqftUnit As Variant, sqFt As Variant, units As Variant
    Dim yearBuilt As Variant, yearRenovated As Variant, appraisedValue As Variant
    Dim noi As Variant, capRate As Variant, locationType As Variant

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Ask user to select Batch folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Batch Folder"
        If .Show <> -1 Then Exit Sub
        batchFolder = .SelectedItems(1)
    End With

    ' Setup filesystem
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(batchFolder)

    ' Create Asset sheet
    On Error Resume Next: Worksheets("Asset").Delete: On Error GoTo 0
    Set assetWS = ThisWorkbook.Sheets.Add
    assetWS.Name = "Asset"

    ' Headers
    With assetWS
        .Range("A1:S1").Value = Array("Loan ID", "Asset ID", "Asset Loan Allocation", "Asset Name", "Asset Address", _
        "Square Footage/Unit", "Square Footage", "Units", "Main Type of Use", "Year Built", "Year Renovate", _
        "Appraised Value", "Appraised Value Date", "Net Operating Income", "Location Type", "Class", _
        "Type of Use Detailed Description", "Cap Rate", "Portfolio")
        .Rows(1).Font.Bold = True
    End With

    rowNum = 2
    assetCounter = 1

    ' Loop through subfolders
    For Each subFolder In folder.SubFolders
        folderName = subFolder.Name
        loanID = Split(folderName, " ")(0) ' Loan ID from subfolder name

        ' Loop through each file in subfolder
        For Each file In subFolder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
                filePath = file.Path
                On Error Resume Next
                Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
                On Error GoTo 0

                If Not wb Is Nothing Then
                    Set ws = Nothing
                    On Error Resume Next: Set ws = wb.Sheets("Cash Flow"): On Error GoTo 0
                    If Not ws Is Nothing Then
                        Dim currentRow As Long, lastRow As Long
                        lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

                        ' Read global values that apply for all properties (if such values are unique)
                        yearBuilt = ""
                        yearRenovated = ""
                        appraisedValue = ""
                        noi = ""
                        locationType = ""
                        capRate = ""
                        sqftUnit = ""
                        sqFt = ""
                        units = ""

                        ' Loop through the rows to find the necessary values
                        For currentRow = 1 To lastRow
                            Select Case LCase(ws.Cells(currentRow, "A").Value)
                                Case "year built"
                                    yearBuilt = ws.Cells(currentRow, "D").Value
                                Case "year rehab"
                                    yearRenovated = ws.Cells(currentRow, "D").Value
                                Case "appraised value"
                                    appraisedValue = ws.Cells(currentRow, "D").Value
                                Case "net operating income"
                                    noi = ws.Cells(currentRow, "L").Value
                                Case "property type"
                                    locationType = ws.Cells(currentRow, "D").Value
                                Case "cap rate"
                                    capRate = ws.Cells(currentRow, "D").Value
                                Case "tot. leasable sq. ft.", "no. units"
                                    sqftUnit = ws.Cells(currentRow, "D").Value
                                    sqFt = sqftUnit
                                    units = sqftUnit
                            End Select
                        Next currentRow

                        ' Now loop again through column A to find every "Property Name" row and create an asset entry for each
                        For currentRow = 1 To lastRow
                            If LCase(ws.Cells(currentRow, "A").Value) = "property name" Then
                                ' Add row to Asset sheet
                                With assetWS
                                    .Cells(rowNum, 1).Value = loanID
                                    .Cells(rowNum, 2).Value = assetID
                                    .Cells(rowNum, 3).Formula = "=IFERROR(L" & rowNum & "/SUMIF($A:$A,$A" & rowNum & ",$L:$L),0)"
                                    .Cells(rowNum, 4).Value = assetName
                                    .Cells(rowNum, 5).Value = "" ' Asset Address - not specified
                                    .Cells(rowNum, 6).Value = ws.Cells(9, 5).Value
                                    .Cells(rowNum, 7).Value = sqFt
                                    .Cells(rowNum, 8).Value = units
                                    .Cells(rowNum, 9).Value = "" ' Main Type of Use - not specified
                                    .Cells(rowNum, 10).Value = yearBuilt
                                    .Cells(rowNum, 11).Value = yearRenovated
                                    .Cells(rowNum, 12).Value = appraisedValue
                                    .Cells(rowNum, 13).Value = "" ' Appraised Value Date - not specified
                                    .Cells(rowNum, 14).Value = noi
                                    .Cells(rowNum, 15).Value = locationType
                                    .Cells(rowNum, 16).Value = "" ' Class - not specified
                                    .Cells(rowNum, 17).Value = "" ' Type of Use Detailed Description - not specified
                                    If IsNumeric(capRate) And capRate <> "" Then
                                        .Cells(rowNum, 18).Value = Format(capRate, "0.00%")
                                    Else
                                        .Cells(rowNum, 18).Value = capRate
                                    End If
                                    .Cells(rowNum, 19).Value = "" ' Portfolio - not specified
                                End With

                                rowNum = rowNum + 1
                                assetCounter = assetCounter + 1
                            End If
                        Next currentRow

                    End If

                    wb.Close SaveChanges:=False
                    Set wb = Nothing
                End If
            End If
        Next file
    Next subFolder

    assetWS.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Asset sheet created with all Property Names and mapped data!", vbInformation
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

