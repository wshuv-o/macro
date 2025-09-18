Function FindAdjacentValueX(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim foundCell As Range
    Dim i As Integer
    Dim checkCell As Range

    ' Clean search text
    searchText = CleanString(searchText)

    ' Step 1: Find the cell with the exact (cleaned) value
    For Each cell In searchRange
        If CleanString(CStr(cell.Value)) = searchText Then
            Set foundCell = cell
            Exit For
        End If
    Next cell

    If foundCell Is Nothing Then
        FindAdjacentValueX = "Not Found"
        Exit Function
    End If

    On Error GoTo CleanExit
    ' Step 2: Get value from the next cell in the specified direction
    If LCase(direction) = "right" Then
        For i = 1 To maxRight
            Set checkCell = ws.Cells(foundCell.Row, foundCell.Column + i)
            If checkCell.MergeCells Then Set checkCell = checkCell.MergeArea.Cells(1, 2)
            If Trim(CStr(checkCell.Value)) <> "" Then
                FindAdjacentValueX = checkCell.Value
                Exit Function
            End If
        Next i

    ElseIf LCase(direction) = "down" Then
        For i = 1 To maxDown
            Set checkCell = ws.Cells(foundCell.Row + i, foundCell.Column)
            If checkCell.MergeCells Then Set checkCell = checkCell.MergeArea.Cells(1, 1)
            If Trim(CStr(checkCell.Value)) <> "" Then
                FindAdjacentValueX = checkCell.Value
                Exit Function
            End If
        Next i
    Else
        FindAdjacentValueX = "Invalid Direction"
        Exit Function
    End If

    FindAdjacentValueX = "No Value Found"
    
CleanExit:
    On Error GoTo 0
End Function
