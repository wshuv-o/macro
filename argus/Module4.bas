Attribute VB_Name = "Module4"
Sub ConvertHierarchyToJSON()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim row As Long
    Dim cellValue As String
    Dim indentLevel As Integer
    Dim jsonArray As String
    Dim columnJson As String
    Dim i As Long
    
    ' Reference to active worksheet
    Set ws = ActiveSheet
    
    ' Find last row and column with data
    'lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Create new worksheet for output
    Set newWs = ws.Parent.Worksheets.Add
    newWs.Name = "JSON_Output_" & Format(Now, "hhmmss")
    
    ' Start building JSON array
    jsonArray = "[" & vbCrLf
    
    ' Process each column starting from column 2
    For col = 2 To lastCol
        columnJson = ConvertColumnToJSON(ws, col, lastRow)
        
        If columnJson <> "" Then
            If col > 2 Then
                jsonArray = jsonArray & "," & vbCrLf
            End If
            jsonArray = jsonArray & columnJson
        End If
    Next col
    
    ' Close JSON array
    jsonArray = jsonArray & vbCrLf & "]"
    
        ' Save JSON to file
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Get the same directory as the Excel file
    filePath = ThisWorkbook.path & "\shuvo.json"
    
    ' Write JSON to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, jsonArray
    Close #fileNum
    
    
    ' Output to new worksheet
    'newWs.Cells(1, 1).value = jsonArray
    'newWs.Columns(1).ColumnWidth = 100
    'newWs.Rows(1).WrapText = True
    
    MsgBox "JSON conversion completed! Check the new worksheet: " & newWs.Name
End Sub

Function ConvertColumnToJSON(ws As Worksheet, colNum As Long, lastRow As Long) As String
    Dim row As Long
    Dim cellValue As String
    Dim indentLevel As Integer
    Dim trimmedValue As String
    Dim jsonResult As String
    Dim hasContent As Boolean
    
    ' Arrays to store the hierarchy
    Dim items() As String
    Dim levels() As Integer
    Dim itemCount As Integer
    
    ' First pass: collect all items and their levels
    itemCount = 0
    ReDim items(0 To 1000)
    ReDim levels(0 To 1000)
    
    For row = 9 To lastRow
        cellValue = ws.Cells(row, colNum).value
        
        If Len(Trim(cellValue)) > 0 Then
            indentLevel = (Len(cellValue) - Len(LTrim(cellValue))) \ 2
            trimmedValue = Trim(cellValue)
            
            ' Clean the value
            trimmedValue = Replace(trimmedValue, """", "")
            trimmedValue = Replace(trimmedValue, Chr(10), "")
            trimmedValue = Replace(trimmedValue, Chr(13), "")
            
            items(itemCount) = trimmedValue
            levels(itemCount) = indentLevel
            itemCount = itemCount + 1
            hasContent = True
            
            If Trim(cellValue) = "Cash Flow Available for Distribution" Then
                Exit For
            End If
        End If
    Next row
    
    If Not hasContent Then
        ConvertColumnToJSON = ""
        Exit Function
    End If
    
    ' Second pass: build JSON
    jsonResult = BuildJSONObject(items, levels, 0, itemCount - 1, 0)
    ConvertColumnToJSON = jsonResult
End Function

Function BuildJSONObject(items() As String, levels() As Integer, startIdx As Integer, endIdx As Integer, baseLevel As Integer) As String
    Dim result As String
    Dim i As Integer
    Dim currentLevel As Integer
    Dim indent As String
    Dim firstProperty As Boolean
    
    result = "{"
    firstProperty = True
    i = startIdx
    
    While i <= endIdx
        currentLevel = levels(i)
        
        ' Only process items at the current base level
        If currentLevel = baseLevel Then
            ' Check if this item has children
            Dim hasChildren As Boolean
            Dim childEndIdx As Integer
            hasChildren = False
            childEndIdx = i
            
            ' Find children and end of this item's scope
            If i < endIdx Then
                Dim j As Integer
                For j = i + 1 To endIdx
                    If levels(j) = currentLevel + 1 And Not hasChildren Then
                        hasChildren = True
                    ElseIf levels(j) <= currentLevel Then
                        childEndIdx = j - 1
                        Exit For
                    Else
                        childEndIdx = j
                    End If
                Next j
            End If
            
            ' Add comma if not first property
            If Not firstProperty Then
                result = result & ","
            End If
            firstProperty = False
            
            ' Add the property
            indent = vbCrLf & String(currentLevel * 2 + 2, " ")
            
            If hasChildren Then
                result = result & indent & """" & items(i) & """: "
                result = result & BuildJSONObject(items, levels, i + 1, childEndIdx, currentLevel + 1)
                i = childEndIdx + 1
            Else
                result = result & indent & """" & items(i) & """: null"
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Wend
    
    result = result & vbCrLf & String(baseLevel * 2, " ") & "}"
    BuildJSONObject = result
End Function
