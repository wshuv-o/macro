Function getValueAtIntersectionString(ws As Worksheet, xHeader As String, yHeader As String, xRange As String, yRange As String) As Variant
    Dim xCell As Range, yCell As Range
    Dim resultCell As Range

    ' Search for xHeader in rows 20 to 30, all columns
    Set xCell = ws.Range(xRange).Find(What:=xHeader, LookIn:=xlValues, LookAt:=xlWhole)
    
    If xCell Is Nothing Then
        MsgBox "X Header Not Found"
        Exit Function
    End If

    ' Search for yHeader in columns A to E, all rows
    Set yCell = ws.Range(yRange).Find(What:=yHeader, LookIn:=xlValues, LookAt:=xlWhole)
    If yCell Is Nothing Then
        MsgBox "Y Header Not Found"
        Exit Function
    End If

    ' Return the intersecting cell's value
    Set resultCell = ws.Cells(yCell.Row, xCell.Column)
    getValueAtIntersectionString = resultCell.Value
End Function
