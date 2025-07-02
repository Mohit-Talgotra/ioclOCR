Private Sub ParseGeminiTableJSON(JsonString As String, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = 1
    
    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    
    pos = 1
    
    Do While pos < Len(JsonString)
        tableStart = InStr(pos, JsonString, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = FindObjectEnd(JsonString, tableStart)
        If tableEnd = 0 Then Exit Do
        
        Dim tableContent As String
        tableContent = Mid(JsonString, tableStart, tableEnd - tableStart + 1)

        Dim headersStart As Long, headersEnd As Long
        headersStart = InStr(tableContent, """headers"":[")
        If headersStart = 0 Then headersStart = InStr(tableContent, """headers"": [")
        
        Dim searchLen As Long
        
        If headersStart > 0 Then
            searchLen = IIf(InStr(tableContent, """headers"":[") > 0, 12, 13)
            headersStart = headersStart + searchLen
            
            headersEnd = FindArrayEnd(tableContent, headersStart)
            If headersEnd > headersStart Then
                Dim headersContent As String
                headersContent = Mid(tableContent, headersStart, headersEnd - headersStart)

                Dim headers As Variant
                headers = ParseJSONStringArray(headersContent)
                
                Dim col As Long
                For col = 0 To UBound(headers)
                    If headers(col) <> "" Then
                        ws.Cells(currentRow, col + 1).Value = headers(col)
                    Else
                        ws.Cells(currentRow, col + 1).Value = "Column " & (col + 1)
                    End If
                    ws.Cells(currentRow, col + 1).Font.Bold = True
                    ws.Cells(currentRow, col + 1).Interior.Color = RGB(220, 220, 220)
                Next col
                currentRow = currentRow + 1
            End If
        End If

        Dim rowsStart As Long, rowsEnd As Long
        rowsStart = InStr(tableContent, """rows"":[")
        If rowsStart = 0 Then rowsStart = InStr(tableContent, """rows"": [")
        
        If rowsStart > 0 Then
            searchLen = IIf(InStr(tableContent, """rows"":[") > 0, 9, 10)
            rowsStart = rowsStart + searchLen
            
            rowsEnd = FindArrayEnd(tableContent, rowsStart)
            If rowsEnd > rowsStart Then
                Dim rowsContent As String
                rowsContent = Mid(tableContent, rowsStart, rowsEnd - rowsStart)
                currentRow = ParseStandardizedTableRows(rowsContent, ws, currentRow)
            End If
        End If

        If currentRow > 1 Then currentRow = currentRow + 1

        pos = tableEnd + 1
    Loop

    If currentRow = 1 Then
        ws.Cells(currentRow, 1).Value = "No table data found in the response"
        ws.Cells(currentRow, 1).Font.Italic = True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing Gemini table JSON: " & Err.Description, vbCritical
End Sub

Private Function ParseStandardizedTableRows(rowsContent As String, ws As Worksheet, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim pos As Long
    pos = 1

    Do While pos < Len(rowsContent)
        pos = InStr(pos, rowsContent, "[")
        If pos = 0 Then Exit Do

        Dim bracketCount As Long
        Dim endPos As Long
        Dim inQuotes As Boolean
        Dim escapeNext As Boolean
        
        bracketCount = 1
        endPos = pos + 1
        inQuotes = False
        escapeNext = False
        
        Do While endPos <= Len(rowsContent) And bracketCount > 0
            Dim currentChar As String
            currentChar = Mid(rowsContent, endPos, 1)
            
            If escapeNext Then
                escapeNext = False
            ElseIf currentChar = "\" Then
                escapeNext = True
            ElseIf currentChar = """" Then
                inQuotes = Not inQuotes
            ElseIf Not inQuotes Then
                If currentChar = "[" Then
                    bracketCount = bracketCount + 1
                ElseIf currentChar = "]" Then
                    bracketCount = bracketCount - 1
                End If
            End If
            endPos = endPos + 1
        Loop
        
        If bracketCount = 0 Then
            Dim rowContent As String
            rowContent = Mid(rowsContent, pos + 1, endPos - pos - 2)
            
            Dim rowData As Variant
            rowData = ParseJSONStringArray(rowContent)

            Dim j As Long
            For j = 0 To UBound(rowData)
                If rowData(j) <> "null" And rowData(j) <> "" Then
                    ws.Cells(currentRow, j + 1).Value = rowData(j)
                End If
            Next j
            
            Call HighlightSignatureDetectedRow(ws, currentRow)
            
            currentRow = currentRow + 1
            pos = endPos
        Else
            Exit Do
        End If
    Loop
    
    ParseStandardizedTableRows = currentRow
    Exit Function
    
ErrorHandler:
    ParseStandardizedTableRows = startRow
End Function

Private Sub PopulateExcelWithGeminiData(jsonData As String)
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Clear existing data in current sheet?", vbYesNoCancel + vbQuestion)
    
    If response = vbCancel Then Exit Sub
    If response = vbYes Then ActiveSheet.Cells.Clear
    
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ParseGeminiTableJSON jsonData, ws

    ws.Columns.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error populating data: " & Err.Description, vbCritical
End Sub

Public Sub ParseGeminiTableJSONToSheet(JsonString As String, ws As Worksheet)
    On Error GoTo ErrorHandler

    ws.Cells.Clear
    
    Dim currentRow As Long
    currentRow = 1

    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    
    pos = 1

    Do While pos < Len(JsonString)
        tableStart = InStr(pos, JsonString, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = FindObjectEnd(JsonString, tableStart)
        If tableEnd = 0 Then Exit Do
        
        Dim tableContent As String
        tableContent = Mid(JsonString, tableStart, tableEnd - tableStart + 1)

        Dim headersStart As Long, headersEnd As Long
        headersStart = InStr(tableContent, """headers"":[")
        If headersStart = 0 Then headersStart = InStr(tableContent, """headers"": [")
        
        Dim searchLen As Long
        
        If headersStart > 0 Then
            searchLen = IIf(InStr(tableContent, """headers"":[") > 0, 12, 13)
            headersStart = headersStart + searchLen
            
            headersEnd = FindArrayEnd(tableContent, headersStart)
            If headersEnd > headersStart Then
                Dim headersContent As String
                headersContent = Mid(tableContent, headersStart, headersEnd - headersStart)

                Dim headers As Variant
                headers = ParseJSONStringArray(headersContent)
                
                Dim col As Long
                For col = 0 To UBound(headers)
                    If headers(col) <> "" Then
                        ws.Cells(currentRow, col + 1).Value = headers(col)
                    Else
                        ws.Cells(currentRow, col + 1).Value = "Column " & (col + 1)
                    End If
                    ws.Cells(currentRow, col + 1).Font.Bold = True
                    ws.Cells(currentRow, col + 1).Interior.Color = RGB(220, 220, 220)
                Next col
                currentRow = currentRow + 1
            End If
        End If

        Dim rowsStart As Long, rowsEnd As Long
        rowsStart = InStr(tableContent, """rows"":[")
        If rowsStart = 0 Then rowsStart = InStr(tableContent, """rows"": [")
        
        If rowsStart > 0 Then
            searchLen = IIf(InStr(tableContent, """rows"":[") > 0, 9, 10)
            rowsStart = rowsStart + searchLen
            
            rowsEnd = FindArrayEnd(tableContent, rowsStart)
            If rowsEnd > rowsStart Then
                Dim rowsContent As String
                rowsContent = Mid(tableContent, rowsStart, rowsEnd - rowsStart)
                currentRow = ParseStandardizedTableRows(rowsContent, ws, currentRow)
            End If
        End If

        If currentRow > 1 Then currentRow = currentRow + 1

        pos = tableEnd + 1
    Loop

    If currentRow = 1 Then
        ws.Cells(currentRow, 1).Value = "No table data found in the response"
        ws.Cells(currentRow, 1).Font.Italic = True
    End If

    ws.Columns.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing Gemini table JSON to sheet: " & Err.Description, vbCritical
End Sub

Private Sub HighlightSignatureDetectedRow(ws As Worksheet, rowNum As Long)

    Dim col As Long
    Dim lastCol As Long
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol

        If InStr(1, ws.Cells(rowNum, col).Value, "SIGNATURE DETECTED", vbTextCompare) > 0 Then

            ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol)).Interior.Color = RGB(198, 239, 206)

            Exit Sub

        End If

    Next col

End Sub

Private Function ParseJSONStringArray(arrayContent As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim items() As String
    Dim itemCount As Long
    Dim i As Long
    Dim inQuotes As Boolean
    Dim currentItem As String
    Dim escapeNext As Boolean
    
    ReDim items(0 To 100)
    itemCount = 0
    currentItem = ""
    inQuotes = False
    escapeNext = False
    
    For i = 1 To Len(arrayContent)
        Dim char As String
        char = Mid(arrayContent, i, 1)
        
        If escapeNext Then
            If char = "n" Then
                currentItem = currentItem & " "
            ElseIf char = "r" Then
                currentItem = currentItem & " "
            ElseIf char = "t" Then
                currentItem = currentItem & " "
            ElseIf char = "\" Then
                currentItem = currentItem & "\"
            ElseIf char = """" Then
                currentItem = currentItem & """"
            Else
                currentItem = currentItem & char
            End If
            escapeNext = False
        ElseIf char = "\" Then
            escapeNext = True
        ElseIf char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            currentItem = Trim(currentItem)
            If Left(currentItem, 1) = """" And Right(currentItem, 1) = """" Then
                currentItem = Mid(currentItem, 2, Len(currentItem) - 2)
            End If
            currentItem = Replace(currentItem, "\\n", " ")
            currentItem = Replace(currentItem, "\\r", " ")
            currentItem = Replace(currentItem, "\\", " ")
            items(itemCount) = currentItem
            itemCount = itemCount + 1
            currentItem = ""
        Else
            If inQuotes Or (char <> " " Or currentItem <> "") Then
                currentItem = currentItem & char
            End If
        End If
    Next i

    If currentItem <> "" Then
        currentItem = Trim(currentItem)
        If Left(currentItem, 1) = """" And Right(currentItem, 1) = """" Then
            currentItem = Mid(currentItem, 2, Len(currentItem) - 2)
        End If
        currentItem = Replace(currentItem, "\\n", " ")
        currentItem = Replace(currentItem, "\\r", " ")
        currentItem = Replace(currentItem, "\\", " ")
        items(itemCount) = currentItem
        itemCount = itemCount + 1
    End If

    ReDim Preserve items(0 To itemCount - 1)
    ParseJSONStringArray = items
    Exit Function
    
ErrorHandler:
    ReDim items(0 To 0)
    items(0) = ""
    ParseJSONStringArray = items
End Function

Private Function FindArrayEnd(text As String, startPos As Long) As Long
    Dim bracketCount As Long
    Dim i As Long
    bracketCount = 1
    
    For i = startPos To Len(text)
        If Mid(text, i, 1) = "[" Then
            bracketCount = bracketCount + 1
        ElseIf Mid(text, i, 1) = "]" Then
            bracketCount = bracketCount - 1
            If bracketCount = 0 Then
                FindArrayEnd = i
                Exit Function
            End If
        End If
    Next i
    
    FindArrayEnd = 0
End Function