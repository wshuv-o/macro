Attribute VB_Name = "Module2"
Sub CopyLineItem()
    Dim ws As Worksheet
    Dim cashFlowSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim col As Long
    
    ' Set the CashFlow sheet
    Set cashFlowSheet = ThisWorkbook.Sheets("CashFlow")
    cashFlowSheet.Cells.ClearContents ' Optional: clear previous contents

    lastRow = 1 ' Start pasting from row 1
    col = 1

    ' Loop through sheets starting from the 3rd
    For i = 3 To ThisWorkbook.Sheets.Count
        Set ws = ThisWorkbook.Sheets(i)
        
        ' Find the last used row in column A of the current sheet
        Dim lastUsedRow As Long
        lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

        ' Copy the range from Column A
        ws.Range("A1:A" & lastUsedRow).Copy
        
        ' Paste into the CashFlow sheet
        cashFlowSheet.Cells(1, col).PasteSpecial Paste:=xlPasteValues
        
        ' Update the lastRow for next paste
        lastRow = cashFlowSheet.Cells(cashFlowSheet.Rows.Count, "A").End(xlUp).row + 2 ' +2 to leave a blank row
        col = col + 1
    Next i

    Application.CutCopyMode = False
    MsgBox "Column A from all sheets (starting from the 3rd) copied to 'CashFlow'."
End Sub

Sub ConvertIndentedRowsToJson()
    Dim ws As Worksheet
    Dim jsonOutput As String
    Dim rowJson As String
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long

    Set ws = ThisWorkbook.Sheets("CashFlow") ' Adjust as needed
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    lastCol = 120 ' As per your specification

    jsonOutput = "["

    ' Loop through each row (record)
    For r = 1 To lastRow
        Dim stack As New Collection
        Dim keyStack As New Collection
        rowJson = ""
        rowJson = rowJson & "{"

        ' Dictionary to store final string at each level
        Dim levelJson(0 To 10) As String
        Dim currentLevel As Integer
        currentLevel = 0

        For c = 1 To lastCol
            Dim rawKey As String
            rawKey = Trim(ws.Cells(r, c).value)

            If rawKey <> "" Then
                Dim indent As Integer
                indent = (Len(ws.Cells(r, c).value) - Len(rawKey)) \ 2 ' 2 spaces = 1 level

                Dim value As String
                value = Trim(ws.Cells(r, c + 1).value)

                ' Close lower levels
                Do While currentLevel > indent
                    rowJson = Left(rowJson, Len(rowJson) - 1) & "},"
                    currentLevel = currentLevel - 1
                Loop

                ' Start new level if needed
                If value = "" Then
                    rowJson = rowJson & """" & rawKey & """: {"
                    currentLevel = indent + 1
                Else
                    rowJson = rowJson & """" & rawKey & """: """ & Replace(value, """", "\""") & ""","
                End If
            End If
        Next c

        ' Close all open brackets
        Do While currentLevel > 0
            rowJson = Left(rowJson, Len(rowJson) - 1) & "},"
            currentLevel = currentLevel - 1
        Loop

        ' Remove last comma and close object
        If Right(rowJson, 1) = "," Then rowJson = Left(rowJson, Len(rowJson) - 1)
        rowJson = rowJson & "},"
        jsonOutput = jsonOutput & vbCrLf & rowJson
    Next r

    ' Finalize JSON array
    If Right(jsonOutput, 1) = "," Then jsonOutput = Left(jsonOutput, Len(jsonOutput) - 1)
    jsonOutput = jsonOutput & vbCrLf & "]"

    ' Output JSON to file
    Dim fso As Object, jsonFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(ThisWorkbook.path & "\ExportedIndentedJson.json", True)
    jsonFile.Write jsonOutput
    jsonFile.Close

    MsgBox "JSON created successfully!"
End Sub

Sub ConvertIndentedColumnsToJSON()
    Dim ws As Worksheet
    Dim jsonOutput As String
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long

    Set ws = ThisWorkbook.Sheets("CashFlow")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    lastCol = 120 ' up to column 120

    jsonOutput = "{"

    ' Loop through each column (each object)
    For c = 1 To lastCol
        Dim colJson As String
        colJson = "{"
        
        Dim levelStack(0 To 10) As Long
        Dim currentLevel As Long: currentLevel = 0
        levelStack(currentLevel) = 0

        For r = 1 To lastRow
            Dim cellVal As String
            cellVal = ws.Cells(r, c).value
            
            If Trim(cellVal) <> "" Then
                Dim rawValue As String: rawValue = LTrim(cellVal)
                Dim indent As Long: indent = (Len(cellVal) - Len(rawValue)) \ 2

                ' Close levels if we're going up
                Do While currentLevel > indent
                    colJson = Left(colJson, Len(colJson) - 1) & "},"
                    currentLevel = currentLevel - 1
                Loop

                ' Increase index for this level
                levelStack(indent) = levelStack(indent) + 1
                Dim key As String: key = CStr(levelStack(indent))

                If indent > currentLevel Then
                    currentLevel = indent
                    levelStack(currentLevel) = 1 ' start count at new level
                End If

                ' Add to JSON
                colJson = colJson & """" & key & """: """
                colJson = colJson & Replace(rawValue, """", "\""") & ""","
            End If
        Next r

        ' Close remaining open levels
        Do While currentLevel > 0
            colJson = Left(colJson, Len(colJson) - 1) & "},"
            currentLevel = currentLevel - 1
        Loop

        If Right(colJson, 1) = "," Then colJson = Left(colJson, Len(colJson) - 1)
        colJson = colJson & "}"
        jsonOutput = jsonOutput & """" & (c - 1) & """: " & colJson & ","
    Next c

    ' Final clean up
    If Right(jsonOutput, 1) = "," Then jsonOutput = Left(jsonOutput, Len(jsonOutput) - 1)
    jsonOutput = jsonOutput & "}"

    ' Write to file
    Dim fso As Object, jsonFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(ThisWorkbook.path & "\FinalNestedJson.json", True)
    jsonFile.Write jsonOutput
    jsonFile.Close

    MsgBox "JSON file created successfully at: " & ThisWorkbook.path
End Sub


Sub ConvertIndentedColumnsToJson2()
    Dim ws As Worksheet
    Dim jsonOutput As String
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long

    Set ws = ThisWorkbook.Sheets("CashFlow")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    lastCol = 120 ' up to column 120

    jsonOutput = "{"

    ' Loop through each column (each object)
    For c = 1 To lastCol
        Dim colJson As String
        colJson = "{"
        
        Dim levelStack(0 To 10) As Long
        Dim currentLevel As Long: currentLevel = 0
        levelStack(currentLevel) = 0

        For r = 1 To lastRow
            Dim cellVal As String
            cellVal = ws.Cells(r, c).value
            
            If Trim(cellVal) <> "" Then
                Dim rawValue As String: rawValue = LTrim(cellVal)
                Dim indent As Long: indent = (Len(cellVal) - Len(rawValue)) \ 2

                ' Close levels if we're going up
                Do While currentLevel > indent
                    colJson = Left(colJson, Len(colJson) - 1) & "},"
                    currentLevel = currentLevel - 1
                Loop

                ' Use rawValue as the key
                Dim key As String: key = rawValue

                ' Add to JSON
                colJson = colJson & """" & Replace(key, """", "\""") & """: """
                colJson = colJson & Replace(rawValue, """", "\""") & ""","
            End If
        Next r

        ' Close remaining open levels
        Do While currentLevel > 0
            colJson = Left(colJson, Len(colJson) - 1) & "},"
            currentLevel = currentLevel - 1
        Loop

        If Right(colJson, 1) = "," Then colJson = Left(colJson, Len(colJson) - 1)
        colJson = colJson & "}"
        jsonOutput = jsonOutput & """" & (c - 1) & """: " & colJson & ","
    Next c

    ' Final clean up
    If Right(jsonOutput, 1) = "," Then jsonOutput = Left(jsonOutput, Len(jsonOutput) - 1)
    jsonOutput = jsonOutput & "}"

    ' Write to file
    Dim fso As Object, jsonFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(ThisWorkbook.path & "\FinalNestedJson.json", True)
    jsonFile.Write jsonOutput
    jsonFile.Close

    MsgBox "JSON file created successfully at: " & ThisWorkbook.path
End Sub


Sub ConvertIndentedColumnsToJson3()
    Dim ws As Worksheet
    Dim jsonOutput As String
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long

    Set ws = ThisWorkbook.Sheets("CashFlow")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    lastCol = 120 ' up to column 120

    jsonOutput = "{"

    ' Loop through each column (each object)
    For c = 1 To lastCol
        Dim colJson As String
        colJson = "{"
        
        Dim levelStack(0 To 10) As Variant
        Dim currentLevel As Long: currentLevel = 0
        levelStack(currentLevel) = ""

        For r = 1 To lastRow
            Dim cellVal As String
            cellVal = ws.Cells(r, c).value
            
            If Trim(cellVal) <> "" Then
                Dim rawValue As String: rawValue = LTrim(cellVal)
                Dim indent As Long: indent = (Len(cellVal) - Len(rawValue)) \ 2

                ' Close levels if we're going up
                Do While currentLevel > indent
                    colJson = Left(colJson, Len(colJson) - 1) & "},"
                    currentLevel = currentLevel - 1
                Loop

                ' Generate key based on hierarchy
                Dim key As String
                If indent > 0 Then
                    Dim parentKey As String
                    If currentLevel > 0 Then
                        parentKey = levelStack(currentLevel - 1)
                    Else
                        parentKey = rawValue
                    End If
                    key = parentKey & "_children"
                Else
                    key = rawValue
                End If

                ' Update level stack
                If indent > currentLevel Then
                    currentLevel = indent
                    levelStack(currentLevel) = rawValue
                Else
                    levelStack(currentLevel) = rawValue
                End If

                ' Add to JSON
                colJson = colJson & """" & Replace(key, """", "\""") & """: """
                colJson = colJson & Replace(rawValue, """", "\""") & ""","
            End If
        Next r

        ' Close remaining open levels
        Do While currentLevel > 0
            colJson = Left(colJson, Len(colJson) - 1) & "},"
            currentLevel = currentLevel - 1
        Loop

        If Right(colJson, 1) = "," Then colJson = Left(colJson, Len(colJson) - 1)
        colJson = colJson & "}"
        jsonOutput = jsonOutput & """" & (c - 1) & """: " & colJson & ","
    Next c

    ' Final clean up
    If Right(jsonOutput, 1) = "," Then jsonOutput = Left(jsonOutput, Len(jsonOutput) - 1)
    jsonOutput = jsonOutput & "}"

    ' Write to file
    Dim fso As Object, jsonFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(ThisWorkbook.path & "\FinalNestedJson.json", True)
    jsonFile.Write jsonOutput
    jsonFile.Close

    MsgBox "JSON file created successfully at: " & ThisWorkbook.path
End Sub



Option Explicit

Sub ConvertIndentedColumnsToJson4()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim jsonOutput As String
    Dim lastRow As Long
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("CashFlow")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row ' Assuming data is in Column 1
    
    Dim stack As Collection
    Set stack = New Collection
    stack.Add CreateObject("Scripting.Dictionary") ' Initialize with root dictionary
    
    ' Process each row (each item)
    For r = 1 To lastRow
        Dim cellVal As String
        cellVal = ws.Cells(r, 2).value
        
        If Trim(cellVal) <> "" Then
            Dim rawValue As String: rawValue = LTrim(cellVal)
            Dim indent As Long: indent = (Len(cellVal) - Len(rawValue)) \ 2
            
            ' Adjust stack to match current indentation level
            While stack.Count > indent + 1
                stack.Remove stack.Count ' Remove last item (mimic stack pop)
            Wend
            
            ' Get parent dictionary
            Dim parentDict As Object
            Set parentDict = stack(stack.Count)
            
            ' Create new dictionary for current level
            Dim currentDict As Object
            Set currentDict = CreateObject("Scripting.Dictionary")
            
            ' Add value to current dictionary
            currentDict.Add rawValue, rawValue
            
            ' Add current dictionary to parent with appropriate key
            Dim parentKey As String
            If stack.Count > 1 Then
                Dim grandParentDict As Object
                Set grandParentDict = stack(stack.Count - 1)
                parentKey = GetParentKey(grandParentDict)
                grandParentDict.Add parentKey & "_children", currentDict
            Else
                parentDict.Add rawValue, currentDict
            End If
            
            ' Push current dictionary to stack
            stack.Add currentDict
        End If
    Next r
    
    ' Build JSON from root dictionary
    jsonOutput = BuildJson(stack(1)) ' Root dictionary is the first item in the stack
    
    ' Write to file
    Dim fso As Object, jsonFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(ThisWorkbook.path & "\FinalNestedJson.json", True)
    jsonFile.Write jsonOutput
    jsonFile.Close
    
    MsgBox "JSON file created successfully at: " & ThisWorkbook.path
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Function GetParentKey(ByVal dict As Object) As String
    Dim key As Variant
    For Each key In dict.Keys
        If Not TypeName(dict(key)) = "Dictionary" Then
            GetParentKey = key
            Exit Function
        End If
    Next key
End Function

Function BuildJson(ByVal dict As Object) As String
    Dim json As String
    json = "{"
    Dim first As Boolean: first = True
    Dim key As Variant
    
    For Each key In dict.Keys
        If Not first Then json = json & ","
        first = False
        
        json = json & """" & Replace(key, """", "\""") & """: "
        
        If TypeName(dict(key)) = "Dictionary" Then
            json = json & BuildJson(dict(key))
        Else
            json = json & """" & Replace(CStr(dict(key)), """", "\""") & """"
        End If
    Next key
    
    json = json & "}"
    BuildJson = json
End Function
