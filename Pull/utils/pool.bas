Function FindAdjacentValueSimilar(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim r As Long, c As Long
    Dim foundCell As Range
    Dim checkCell As Range

    ' Search for the searchText (partial match using Like)
    For Each cell In searchRange
        If Not IsError(cell.Value) Then
            If Trim(cell.Value) Like "*" & searchText & "*" Then
                Set foundCell = cell
                Exit For
            End If
        End If
    Next cell

    If foundCell Is Nothing Then
        FindAdjacentValueSimilar = "Not Found"
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
                    FindAdjacentValueSimilar = checkCell.Value
                    Exit Function
                End If
            Else
                Set checkCell = checkCell.MergeArea.Cells(1, 1)
                If Trim(checkCell.Value) <> "" Then
                    FindAdjacentValueSimilar = checkCell.Value
                    Exit Function
                End If
            End If
            On Error GoTo 0
        Next i
        FindAdjacentValueSimilar = "No Value Found"

    ElseIf direction = "down" Then
        Dim j As Integer
        For j = 1 To maxDown
            Set checkCell = ws.Cells(r + j, c)
            If Trim(checkCell.Value) <> "" Then
                FindAdjacentValueSimilar = checkCell.Value
                Exit Function
            End If
        Next j
        FindAdjacentValueSimilar = "No Value Found"
    Else
        FindAdjacentValueSimilar = "Invalid Direction"
    End If
End Function


Function FindAdjacentValueWS(ws As Worksheet, searchText As String, direction As String, searchRange As Range, maxRight As Integer, maxDown As Integer) As Variant
    Dim cell As Range
    Dim r As Long, c As Long
    Dim foundCell As Range
    Dim checkCell As Range
    
    ' Search for the searchText in the given range
'    For Each cell In searchRange
 '       If cell.Value = searchText Then
  '          Set foundCell = cell
     '       Exit For
   '     End If
    'Next cell
    For Each cell In searchRange
    If Not IsError(cell.Value) Then
        If Trim(cell.Value) = searchText Then
            Set foundCell = cell
            Exit For
        End If
    End If
Next cell

    If foundCell Is Nothing Then
        FindAdjacentValueWS = "Not Found"
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
                    FindAdjacentValueWS = checkCell.Value
                    Exit Function
                End If
            Else
                Set checkCell = checkCell.MergeArea.Cells(1, 1)
                If Trim(checkCell.Value) <> "" Then
                    FindAdjacentValueWS = checkCell.Value
                    Exit Function
                End If
            End If
            On Error GoTo 0
        Next i
        FindAdjacentValueWS = "No Value Found"
        
    ElseIf direction = "down" Then
        Dim j As Integer
        For j = 1 To maxDown
            Set checkCell = ws.Cells(r + j, c)
            If Trim(checkCell.Value) <> "" Then
                FindAdjacentValueWS = checkCell.Value
                Exit Function
            End If
        Next j
        FindAdjacentValueWS = "No Value Found"
    Else
        FindAdjacentValueWS = "Invalid Direction"
    End If
End Function

Function getValueAtIntersection(ws As Worksheet) As Variant
    Dim xCell As Range, yCell As Range
    Dim resultCell As Range

    Dim xHeader As String
    Dim yHeader As String
    xHeader = "Underwritten"
    yHeader = "Debt Service on Recommended loan"

    ' Search for xHeader in rows 20 to 30, all columns
    Set xCell = ws.Range("A20:AP30").Find(What:=xHeader, LookIn:=xlValues, LookAt:=xlWhole)
    
    If xCell Is Nothing Then
        MsgBox "X Header Not Found"
        Exit Function
    End If

    ' Search for yHeader in columns A to E, all rows
    Set yCell = ws.Range("A1:E100").Find(What:=yHeader, LookIn:=xlValues, LookAt:=xlWhole)
    If yCell Is Nothing Then
        MsgBox "Y Header Not Found"
        Exit Function
    End If

    ' Return the intersecting cell's value
    Set resultCell = ws.Cells(yCell.Row, xCell.Column)
    getValueAtIntersection = resultCell.Value
End Function
Function getValueAtIntersectionString(ws As Worksheet, xHeader As String, yHeader As String, xRange As String, yRange As String) As Variant
    Dim xCell As Range, yCell As Range
    Dim resultCell As Range

    ' Search for xHeader in rows 20 to 30, all columns
    Set xCell = ws.Range(xRange).Find(What:=xHeader, LookIn:=xlValues, LookAt:=xlWhole)
    
    If xCell Is Nothing Then
        'MsgBox "X Header Not Found"
        Exit Function
    End If

    ' Search for yHeader in columns A to E, all rows
    Set yCell = ws.Range(yRange).Find(What:=yHeader, LookIn:=xlValues, LookAt:=xlWhole)
    If yCell Is Nothing Then
        'MsgBox "Y Header Not Found"
        Exit Function
    End If

    ' Return the intersecting cell's value
    Set resultCell = ws.Cells(yCell.Row, xCell.Column)
    getValueAtIntersectionString = resultCell.Value
End Function

Function FirstNonEmpty(ws As Worksheet, rows As Variant, col As String, dataRng As String, hdrRng As String) As Variant
    Dim r As Variant, val As Variant
    For Each r In rows
        On Error Resume Next
        val = getValueAtIntersectionString(ws, CStr(r), col, dataRng, hdrRng)
        On Error GoTo 0
        If Not IsError(val) Then
            If Trim(val & "") <> "" Then
                FirstNonEmpty = val
                Exit Function
            End If
        End If
    Next
    FirstNonEmpty = "No Value Found"
End Function
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
