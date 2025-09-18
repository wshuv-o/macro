Function GetRange(ws As Worksheet, searchText As String, startAddress As String, maxRight As Integer, maxDown As Integer) As Range
    Dim startCell As Range
    Dim r As Long, c As Long
    Dim currentCell As Range
    
    Set startCell = ws.Range(startAddress)

    For r = 0 To maxDown
        For c = 0 To maxRight
            Set currentCell = startCell.Offset(r, c)
            If Not IsError(currentCell.Value) Then
                If Trim(CStr(currentCell.Value)) = Trim(searchText) Then
                    Set GetRange = currentCell ' Fixed: Use Set for Range assignment
                    Exit Function
                End If
            End If
        Next c
    Next r
    
    Set GetRange = Nothing ' Return Nothing if not found
End Function
