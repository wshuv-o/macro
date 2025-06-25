Attribute VB_Name = "Module5"

Option Explicit

Sub ConvertToJsonGrok()
    Dim ws As Worksheet, newWs As Worksheet
    Dim col As Long, row As Long, lastCol As Long, lastRow As Long
    Dim jsonArray As String, jsonCount As Long
    Dim stack() As String, stackSize As Long
    Dim cellValue As String, indentLevel As Long
    Dim outputLine As String
    Dim i As String
    ' Get active worksheet and find last column
    Set ws = ActiveSheet
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Create new worksheet for output
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("JSON Output").Delete
    Application.DisplayAlerts = True
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "JSON Output"
    On Error GoTo 0
    
    ' Initialize JSON array
    ReDim jsonArray(0 To lastCol - 2)
    jsonCount = 0
    
    ' Process each column starting from 2nd column
    For col = 2 To lastCol
        lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).row
        If lastRow > 1 Then
            ' Initialize stack for hierarchy tracking
            ReDim stack(0 To lastRow)
            stackSize = 0
            outputLine = "{"
            
            Dim firstItem As Boolean
            firstItem = True
            
            ' Process each non-empty cell in column
            For row = 2 To lastRow
                cellValue = Trim(ws.Cells(row, col).value)
                If Len(cellValue) > 0 Then
                    ' Calculate indentation level (2 spaces per level)
                    indentLevel = (Len(cellValue) - Len(LTrim(cellValue))) \ 2
                    
                    ' Pop stack to match current indentation level
                    While stackSize > indentLevel + 1
                        stackSize = stackSize - 1
                        outputLine = outputLine & String(stackSize * 2, " ") & "}"
                        If stackSize > indentLevel + 1 Then
                            outputLine = outputLine & ","
                        End If
                    Wend
                    
                    ' Clean item name and add to output
                    cellValue = Trim(cellValue)
                    cellValue = Replace(cellValue, """", "\""")
                    
                    If firstItem Then
                        firstItem = False
                    Else
                        outputLine = outputLine & ","
                    End If
                    
                    outputLine = outputLine & vbCrLf & String((indentLevel + 1) * 2, " ") & """" & cellValue & """: "
                    
                    ' Check if next row is a child (deeper indentation)
                    Dim nextIndent As Long
                    If row < lastRow Then
                        Dim nextValue As String
                        nextValue = Trim(ws.Cells(row + 1, col).value)
                        If Len(nextValue) > 0 Then
                            nextIndent = (Len(nextValue) - Len(LTrim(nextValue))) \ 2
                        Else
                            nextIndent = 0
                        End If
                    Else
                        nextIndent = 0
                    End If
                    
                    If nextIndent > indentLevel Then
                        outputLine = outputLine & "{"
                        stackSize = stackSize + 1
                        stack(stackSize) = cellValue
                    Else
                        outputLine = outputLine & "{}"
                    End If
                End If
            Next row
            
            ' Close remaining open braces
            While stackSize > 0
                stackSize = stackSize - 1
                outputLine = outputLine & vbCrLf & String(stackSize * 2, " ") & "}"
            Wend
            
            ' Add to JSON array
            jsonArray(jsonCount) = outputLine
            jsonCount = jsonCount + 1
        End If
    Next col
    
    ' Write JSON array to new worksheet
    If jsonCount > 0 Then
        ReDim Preserve jsonArray(0 To jsonCount - 1)
        newWs.Cells(1, 1).value = "[" & vbCrLf & Join(jsonArray, "," & vbCrLf) & vbCrLf & "]"
    Else
        newWs.Cells(1, 1).value = "[]"
    End If
End Sub

