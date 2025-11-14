Function FirstNonEmptyAdjacent(ws As Worksheet, searchTextArray As Variant, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim r As Variant, val As Variant
    For Each r In searchTextArray
        On Error Resume Next
        val = FindAdjacentValueWS(ws, CStr(r), "right", searchRange, maxRight, maxDown)
        On Error GoTo 0
        If Not IsError(val) Then
            If Trim(val & "") <> "" And val <> "Not Found" Then
                FirstNonEmptyAdjacent = val
                Exit Function
            End If
        End If
    Next
    FirstNonEmptyAdjacent = "No Value Found"
End Function
'Add description here how to use this function
