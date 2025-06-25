Attribute VB_Name = "Module3"
Sub ConvertIndentedColumnsToJSON()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim line As String, indentLevel As Long
    Dim stack() As String
    Dim json As String
    Dim i As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    json = "{"
    ReDim stack(0)
    
    For c = 2 To lastCol
        lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).row
        
        For r = 1 To lastRow
            line = ws.Cells(r, c).value
            If Trim(line) <> "" Then
                indentLevel = (Len(line) - Len(LTrim(line))) \ 2 ' 2 spaces per level assumed
                
                ' Close deeper levels
                Do While indentLevel < UBound(stack)
                    json = Left(json, Len(json) - 1) & vbNewLine
                    json = json & String(2 * (UBound(stack) - 1), " ") & "}," & vbNewLine
                    ReDim Preserve stack(UBound(stack) - 1)
                Loop
                
                ' Add comma if not first entry at this level
                If Right(Trim(json), 1) <> "{" And Right(Trim(json), 1) <> "[" Then
                    json = Left(json, Len(json) - 1) & "," & vbNewLine
                End If
                
                ' Add key
                json = json & String(2 * indentLevel, " ") & """" & Trim(line) & """: "
                json = json & "{" & vbNewLine
                
                ' Update stack
                ReDim Preserve stack(indentLevel + 1)
                stack(indentLevel) = Trim(line)
            End If
        Next r
    Next c
    
    ' Close all remaining levels
    For i = UBound(stack) - 1 To 0 Step -1
        json = Left(json, Len(json) - 1) & vbNewLine
        json = json & String(2 * i, " ") & "}," & vbNewLine
    Next i
    
    json = Left(json, Len(json) - 3) & vbNewLine & "}"
    
    ' Output to new worksheet
    Dim outSheet As Worksheet
    Set outSheet = Worksheets.Add
    outSheet.Name = "JSON_Output"
    outSheet.Range("A1").value = json
    
    MsgBox "JSON conversion complete. See 'JSON_Output' sheet."
End Sub


Sub ConvertIndentedColumnsToValidJSONMAYBE()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim line As String, indentLevel As Long
    Dim jsonLines() As String
    Dim lineCount As Long
    lineCount = 0

    ReDim jsonLines(0 To 100000) ' Preallocate space

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim currentIndent As Long
    Dim prevIndent As Long
    Dim firstItem As Boolean
    firstItem = True

    jsonLines(lineCount) = "{"
    lineCount = lineCount + 1

    For c = 2 To lastCol
        lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).row

        For r = 1 To lastRow
            line = ws.Cells(r, c).value
            If Trim(line) <> "" Then
                currentIndent = (Len(line) - Len(LTrim(line))) \ 2

                ' Close deeper levels
                Do While currentIndent < prevIndent
                    jsonLines(lineCount) = Space(2 * prevIndent) & "}"
                    lineCount = lineCount + 1
                    prevIndent = prevIndent - 1
                Loop

                ' Add comma to previous line if not the first or a brace
                If Not firstItem Then
                    Dim lastLine As String
                    lastLine = jsonLines(lineCount - 1)
                    If Right(Trim(lastLine), 1) <> "{" And Right(Trim(lastLine), 1) <> "," Then
                        jsonLines(lineCount - 1) = lastLine & ","
                    End If
                End If

                ' Add current key
                jsonLines(lineCount) = Space(2 * currentIndent) & """" & Trim(line) & """: {"
                lineCount = lineCount + 1

                prevIndent = currentIndent
                firstItem = False
            End If
        Next r
    Next c

    ' Close remaining levels
    Do While prevIndent >= 0
        jsonLines(lineCount) = Space(2 * prevIndent) & "}"
        lineCount = lineCount + 1
        prevIndent = prevIndent - 1
    Loop

    ' Output to a new sheet
    Dim outSheet As Worksheet
    Set outSheet = Worksheets.Add
    outSheet.Name = "Valid_JSON"

    Dim i As Long
    For i = 0 To lineCount - 1
        outSheet.Cells(i + 1, 1).value = jsonLines(i)
    Next i

    MsgBox "Valid JSON created in 'Valid_JSON' sheet."
End Sub

Sub ConvertEachColumnToJSONObject1()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastCol As Long, lastRow As Long
    Dim col As Long, row As Long
    Dim cellValue As String
    Dim currentIndent As Long, prevIndent As Long
    Dim output() As String
    Dim lineCount As Long: lineCount = 0
    Dim colStart As Long: colStart = 2
    
    ReDim output(0 To 100000) ' Preallocate enough space
    
    output(lineCount) = "["
    lineCount = lineCount + 1
    
    For col = colStart To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        output(lineCount) = "  {"
        lineCount = lineCount + 1
        
        prevIndent = -1
        
        lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).row
        For row = 1 To lastRow
            cellValue = ws.Cells(row, col).value
            If Trim(cellValue) <> "" Then
                currentIndent = (Len(cellValue) - Len(LTrim(cellValue))) \ 2
                
                ' Close deeper levels
                Do While currentIndent <= prevIndent
                    output(lineCount - 1) = RTrim(output(lineCount - 1)) ' Remove trailing comma
                    output(lineCount) = Space((prevIndent + 1) * 2) & "}"
                    lineCount = lineCount + 1
                    prevIndent = prevIndent - 1
                Loop
                
                output(lineCount) = Space((currentIndent + 1) * 2) & """" & Trim(cellValue) & """: {"
                lineCount = lineCount + 1
                prevIndent = currentIndent
            End If
        Next row
        
        ' Close remaining levels
        Do While prevIndent >= 0
            output(lineCount - 1) = RTrim(output(lineCount - 1)) ' Remove trailing comma
            output(lineCount) = Space((prevIndent + 1) * 2) & "}"
            lineCount = lineCount + 1
            prevIndent = prevIndent - 1
        Loop
        
        ' Close column object
        output(lineCount) = "  },"
        lineCount = lineCount + 1
    Next col
    
    ' Remove final comma and close array
    output(lineCount - 1) = "  }"
    output(lineCount) = "]"
    
    ' Output to a new worksheet
    Dim outSheet As Worksheet
    Set outSheet = Worksheets.Add
    outSheet.Name = "JSON_Output12"
    
    Dim i As Long
    For i = 0 To lineCount
        outSheet.Cells(i + 1, 1).value = output(i)
    Next i
    
    MsgBox "Finished! JSON array created in 'JSON_Output'."
End Sub

Option Explicit

Sub Cohere()
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim colIndex As Long, rowIndex As Long
    Dim jsonArray As String, jsonObject As String
    Dim indentLevel As Long, prevIndentLevel As Long
    Dim cellValue As String, indentCount As Long
    Dim stack As Object, stackItem As Object
    Dim outputRow As Long

    ' Initialize
    Set ws = ThisWorkbook.ActiveSheet
    Set stack = CreateObject("System.Collections.Stack")
    jsonArray = "["
    outputRow = 1

    ' Find last row and column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Process each column starting from the 2nd column
    For colIndex = 2 To lastCol
        jsonObject = "{"
        indentLevel = -1
        prevIndentLevel = -1

        ' Reset stack for each column
        Set stack = CreateObject("System.Collections.Stack")

        ' Process each row in the column
        For rowIndex = 1 To lastRow
            cellValue = Trim(ws.Cells(rowIndex, colIndex).value)
            If cellValue = "" Then Continue For

            ' Calculate indentation level
            indentCount = Len(cellValue) - Len(Trim(cellValue))
            indentLevel = indentCount / 2

            ' Adjust stack based on indentation
            While stack.Count > 0 And indentLevel <= stack.Peek.Level
                jsonObject = jsonObject & "}"
                stack.Pop
            Wend

            ' Add new key to JSON
            If stack.Count = 0 Then
                jsonObject = jsonObject & """" & cellValue & """: "
            Else
                jsonObject = jsonObject & """, """ & cellValue & """: "
            End If

            ' Add empty object and push to stack
            jsonObject = jsonObject & "{}"
            Set stackItem = New stackItem
            stackItem.Level = indentLevel
            stack.Push stackItem

            ' Update previous indentation level
            prevIndentLevel = indentLevel
        Next rowIndex

        ' Close remaining objects
        While stack.Count > 0
            jsonObject = jsonObject & "}"
            stack.Pop
        Wend

        ' Add column's JSON to array
        If colIndex > 2 Then jsonArray = jsonArray & ", "
        jsonArray = jsonArray & jsonObject
    Next colIndex

    ' Close JSON array
    jsonArray = jsonArray & "]"

    ' Output to new worksheet
    Set newWs = ThisWorkbook.Worksheets.Add
    newWs.Cells(1, 1).value = jsonArray
End Sub

Private Class StackItem
    Public Level As Long
End Class

