Function FindAdjacentValue(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim r As Long, c As Long
    Dim foundCell As Range
    Dim checkCell As Range
    
    ' Search for the searchText in the given range
    For Each cell In searchRange
        If cell.Value = searchText Then
            Set foundCell = cell
            Exit For
        End If
    Next cell
    
    If foundCell Is Nothing Then
        FindAdjacentValue = "Not Found"
        Exit Function
    End If
    
    r = foundCell.Row
    c = foundCell.Column
    
    If direction = "right" Then
        Dim i As Integer
        For i = 1 To maxRight
            On Error Resume Next
            Set checkCell = ws.Cells(r, c + i)
            If Not checkCell.MergeCells Then
                If Trim(checkCell.Value) <> "" Then
                    FindAdjacentValue = checkCell.Value
                    Exit Function
                End If
            Else
                Set checkCell = checkCell.MergeArea.Cells(1, 1)
                If Trim(checkCell.Value) <> "" Then
                    FindAdjacentValue = checkCell.Value
                    Exit Function
                End If
            End If
            On Error GoTo 0
        Next i
        FindAdjacentValue = "No Value Found"
        
    ElseIf direction = "down" Then
        Dim j As Integer
        For j = 1 To maxDown
            Set checkCell = ws.Cells(r + j, c)
            If Trim(checkCell.Value) <> "" Then
                FindAdjacentValue = checkCell.Value
                Exit Function
            End If
        Next j
        FindAdjacentValue = "No Value Found"
    Else
        FindAdjacentValue = "Invalid Direction"
    End If
End Function



'result = FindAdjacentValue(Sheet1, "Name", "right", Sheet1.Range("A1:D20"), 5, 5)

