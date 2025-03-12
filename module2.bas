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
