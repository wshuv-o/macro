Attribute VB_Name = "Module1"
Sub cash_flow()
Dim wb As Workbook
Dim path As String
Dim currentPath As String
Dim dr As String
Dim arr() As Variant
Dim i As Integer
arr = Array("CITI2025001 - Aurora Marketplace", "CITI2025003 - 46th & 48th Portfolio", "CITI2025006 - Warwick Denver", "CITI2025007 - Marysville Shopping Center", "CITI2025010 - Hotel Hendricks", "CITI2025013 - The One at Fayetteville", "CITI2025014 - The Atrium", "CITI2025015 - Columbus Commons", "CITI2025020 - Excelsior on the Park", "CITI2025021 - Cottage Grove", "CITI2025022 - Canton Club East Apartments", "CITI2025023 - Redmond Town Center", "CITI2025025 - Bar Harbour", "CITI2025028 - The Wave", "CITI2025029 - WallyPark SeaTAC", "CITI2025032 -Doubletree Hilton Binghamton", "CITI2025033 - Asden Portfolio", "CITI2025034 - Maple Leaf Apartments", "CITI2025039 - 655 3rd Ave")
For i = 0 To (UBound(arr) - LBound(arr) + 1)
path = ThisWorkbook.path
currentPath = Left(path, InStrRev(path, "\") - 1)
dr = Dir(currentPath & "\" & arr(i) & "\UW**UW*.xls*")
Do While dr <> ""
fullPath = currentPath & "\" & arr(i) & "\" + dr
dr = Dir()
Set wb = Workbooks.Open(fullPath)
wb.Worksheets("Cash Flow").Copy After:=ThisWorkbook.Worksheets(1)
wb.Close SaveChanges:=False
If ActiveSheet.Range("E6").Value <> "" Then
ActiveSheet.Name = ActiveSheet.Range("E6").Value
ElseIf ActiveSheet.Range("C3").Value <> "" Then
ActiveSheet.Name = ActiveSheet.Range("C3").Value
ElseIf ActiveSheet.Range("D5").Value <> "" Then
ActiveSheet.Name = ActiveSheet.Range("D5").Value
End If
Loop
Next i
End Sub

Sub cop()
Dim wb As Workbook
Dim path As String
Dim dr As String
Dim fullPath As String
Dim arr() As Variant
Dim i As Integer
arr = Array("CITI2025005 - Prime 15 Portfolio", "CITI2025033 - Asden Portfolio")
For i = 0 To (UBound(arr) - LBound(arr) + 1)
path = ThisWorkbook.path
currentPath = Left(path, InStrRev(path, "\") - 1)
dr = Dir(currentPath & "\" & arr(i) & "\UW**UW*.xls*")
fullPath = currentPath & "\" & arr(i) & "\" + dr
Set wb = Workbooks.Open(fullPath)
Application.DisplayAlerts = False
For Each ws In wb.Worksheets
If ws.Name Like "Cash Flow*" Then
ws.Copy After:=ThisWorkbook.Worksheets(1)
End If
Next ws
wb.Close False
If ActiveSheet.Range("C3").Value <> "" Then
ActiveSheet.Name = ActiveSheet.Range("C3").Value
ElseIf ActiveSheet.Range("D5").Value <> "" Then
ActiveSheet.Name = ActiveSheet.Range("D5").Value
End If
Application.DisplayAlerts = True
Next i
End Sub

