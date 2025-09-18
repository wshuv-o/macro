Option Explicit

Sub ListSheetsFromFilesFast()
    Dim startFolder As String
    Dim prefix As String
    Dim fso As Object, folder As Object
    Dim rowOut As Long
    Dim wsOut As Worksheet
    
    ' Use the active sheet
    Set wsOut = ActiveSheet
    
    ' Ask user for folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Search"
        If .Show <> -1 Then Exit Sub
        startFolder = .SelectedItems(1)
    End With
    
    ' Ask user for file prefix
    prefix = InputBox("Enter file prefix (e.g., 'uw')", "File Prefix")
    If prefix = "" Then Exit Sub
    
    ' Clear existing content in the active sheet
    wsOut.Cells.Clear
    wsOut.Range("A1").Value = "File Path"
    wsOut.Range("B1").Value = "Sheet Names"
    rowOut = 2
    
    ' Speed optimizations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' File system object for recursive search
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(startFolder)
    
    ' Recursive search
    Call SearchFolderFast(folder, prefix, rowOut, wsOut)
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Done! Found and listed sheet names.", vbInformation
End Sub

Private Sub SearchFolderFast(ByVal folder As Object, ByVal prefix As String, ByRef rowOut As Long, wsOut As Worksheet)
    Dim file As Object, subFolder As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim col As Long
    
    ' Loop through files
    For Each file In folder.Files
        If (LCase(file.Name) Like LCase(prefix & "*.xls*")) Then
            On Error Resume Next
            Set wb = Workbooks.Open(file.Path, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
            If Not wb Is Nothing Then
                ' Write file path
                wsOut.Cells(rowOut, 1).Value = file.Path
                col = 2
                ' List sheet names
                For Each ws In wb.Worksheets
                    wsOut.Cells(rowOut, col).Value = ws.Name
                    col = col + 1
                Next ws
                wb.Close SaveChanges:=False
                rowOut = rowOut + 1
            End If
            On Error GoTo 0
        End If
    Next file
    
    ' Loop through subfolders
    For Each subFolder In folder.SubFolders
        Call SearchFolderFast(subFolder, prefix, rowOut, wsOut)
    Next subFolder
End Sub

