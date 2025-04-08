VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Reset Form"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    ' remove cashflow

End Sub

Private Sub CheckBox5_Click()
    'Remove all the Rent roll tab
End Sub

Private Sub CheckBox2_Click()
    ' If CheckBox2 is checked, reset the value of other checkboxes to False

End Sub

Private Sub CheckBox3_Click()
    ' If CheckBox3 is checked, reset the value of other checkboxes to False

End Sub

Private Sub CheckBox4_Click()
    ' Select all checkboxes when CheckBox4 is checked
    If CheckBox4.value = True Then
        CheckBox1.value = True
        CheckBox2.value = True
        CheckBox3.value = True
        CheckBox5.value = True
        CheckBox6.value = True
        CheckBox7.value = True
        CheckBox8.value = True
        CheckBox9.value = True
    Else
        CheckBox1.value = False
        CheckBox2.value = False
        CheckBox3.value = False
        CheckBox5.value = False
        CheckBox6.value = False
        CheckBox7.value = False
        CheckBox8.value = False
        CheckBox9.value = False
    End If
End Sub


Private Sub CheckBox6_Click()

End Sub

Private Sub CheckBox7_Click()

End Sub

Private Sub CheckBox8_Click()

End Sub

Private Sub CommandButton1_Click()
    ' Cancel - Close the form without doing anything
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    ' Reset - Perform the actions based on user selection and close the form
    
    Dim ws As Worksheet
    Dim mappingSheet As Worksheet
    Dim deleteRows As Range
    Dim lastRow As Long
    Dim i As Long
    
    ' Delete all sheets with "Cash Flow" in A1 if CheckBox1 is selected
    If CheckBox1.value Then
        For Each ws In ThisWorkbook.Sheets
            If ws.Range("A1").value = "Cash Flow" Then
                Application.DisplayAlerts = False ' Prevent confirmation prompt
                ws.Delete
                Application.DisplayAlerts = True
            End If
        Next ws
    End If
    
    ' Reset "Tracker" tab (remove rows starting from row 2) if CheckBox2 is selected
    If CheckBox2.value Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Tracker")
        Set mappingSheet = ThisWorkbook.Sheets("Mapping")
        mappingSheet.Range("D5:D9").ClearContents
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow > 1 Then
                ws.Rows("2:" & lastRow).Delete
            End If
        End If
    End If

    ' Reset "Financials" tab (remove rows starting from row 4) if CheckBox3 is selected
    If CheckBox3.value Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Financials")
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 4 Then
                ws.Rows("4:" & lastRow).Delete
            End If
        End If
    End If
    
    ' Delete all sheets with "Rent Roll Analysis" in A1 if CheckBox5 is selected
    If CheckBox5.value Then
        For Each ws In ThisWorkbook.Sheets
            If ws.Range("A1").value = "Rent Roll Analysis" Then
                Application.DisplayAlerts = False ' Prevent confirmation prompt
                ws.Delete
                Application.DisplayAlerts = True
            End If
        Next ws
    End If

    
    ' Reset "Rent Roll" tab (remove rows starting from row 3) if CheckBox6 is selected
    If CheckBox6.value Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Rent Roll")
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow > 3 Then
                ws.Rows("3:" & lastRow).Delete
            End If
        End If
    End If

    ' Reset "MF Rent Rolls" tab (remove rows starting from row 2) if CheckBox7 is selected
    If CheckBox7.value Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("MF Rent Rolls")
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 2 Then
                ws.Rows("2:" & lastRow).Delete
            End If
        End If
    End If
    
    ' Reset "Loan" tab (remove rows starting from row 6) if CheckBox8 is selected
    If CheckBox8.value Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Loan")
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
            If lastRow >= 6 Then
                ws.Rows("6:" & lastRow).Delete
            End If
        End If
    End If


    ' Reset "Asset" tab (remove rows starting from row 6) if CheckBox7 is selected
    If CheckBox9.value Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Asset")
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 6 Then
                ws.Rows("6:" & lastRow).Delete
            End If
        End If
    End If


    ' Close the form after performing actions
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
