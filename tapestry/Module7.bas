Attribute VB_Name = "Module7"
Public DisableSave As Boolean  ' Global variable to track save blocking

Sub RemoveTemplate()
    Dim ws As Worksheet
    Dim UserResponse As VbMsgBoxResult
    Dim vbComp As Object
    Dim VBProj As Object

    ' Confirmation message
    UserResponse = MsgBox("Are you sure you want to remove the following tabs and all macros?" & _
        Chr(10) & " " & _
        Chr(10) & "O UW File Name" & _
        Chr(10) & "O Main" & _
        Chr(10) & "O All the macros associated with this workbook" & _
        Chr(10) & " " & _
        Chr(10) & "If you proceed, the save operation will not work. You have to 'Save As' the file.", _
        vbYesNo + vbExclamation, "Confirm Action")

    ' If the user clicks "No", exit
    If UserResponse <> vbYes Then Exit Sub

    ' Enable save blocking
    DisableSave = True

    ' Delete "Main" sheet if it exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Main")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    ' Delete "UW File Name" sheet if it exists
    Set ws = ThisWorkbook.Sheets("UW File Name")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Remove all macros (VBA modules)
    On Error Resume Next
    Set VBProj = ThisWorkbook.VBProject
    For Each vbComp In VBProj.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then  ' Modules, Class Modules, UserForms
            VBProj.VBComponents.Remove vbComp
        End If
    Next vbComp
    On Error GoTo 0

    ' Notify user
    MsgBox "Main and UW File Name sheets have been removed, along with all macros. You must use 'Save As' to save this file.", _
        vbInformation, "Process Complete"
End Sub


Sub EnableSave()
    DisableSave = False
    MsgBox "Save functionality has been restored.", vbInformation, "Save Enabled"
End Sub




