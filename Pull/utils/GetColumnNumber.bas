Function GetColumnNumber(ws As Worksheet, searchValue As String, rowNumber As Long) As Long
    Dim rng As Range
    
    ' Search for the string in the specified row
    Set rng = ws.rows(rowNumber).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Return column number if found, else return 0
    If Not rng Is Nothing Then
        GetColumnNumber = rng.Column
    Else
        GetColumnNumber = 5
    End If
End Function
