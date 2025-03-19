VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4725
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public countValue As Integer
Public previousBlock As Integer
Public value As Integer

' Initializes the form with the total value (file count)
Public Sub InitializeForm(initialValue As Integer)
    value = initialValue
    countValue = 0 ' Initialize counter
    previousBlock = 0 ' Initialize previous block for progress bar
    Me.lblCounter.Caption = "0%" ' Set initial caption for counter
    Me.lblProgressBar.Caption = String(0, ChrW(9632)) ' Set initial progress bar to empty
End Sub

' Updates the counter and progress bar
Public Sub UpdateProgress()
    Dim percentage As Double
    percentage = (countValue / value) * 100 ' Calculate the percentage of completion

    ' Update the label with the current percentage
    Me.lblCounter.Caption = Int(percentage) & "%"

    ' Update progress bar every 5%
    If Int(percentage / 5) > previousBlock Then
        previousBlock = Int(percentage / 5)
        Me.lblProgressBar.Caption = String(previousBlock, ChrW(9632)) ' Update progress bar
    End If

    ' Close the form if 100% is reached
    If percentage >= 100 Then
        Me.lblCounter.Caption = "100%" ' Ensure counter shows 100%
        Me.lblProgressBar.Caption = String(20, ChrW(9632)) ' Fill the progress bar
        Me.CloseF ' Close the form
    End If
End Sub

' Increments the count value and schedules next update
Public Sub IncrementCount()
    countValue = countValue + 1 ' Increment the counter
End Sub

' Close the form
Public Sub CloseF()
    Unload Me
End Sub


Private Sub UserForm_Click()

End Sub
