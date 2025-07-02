Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim protectedNames As Variant
    protectedNames = Array("Dashboard", "Summary", "Charts")

    Dim nameIsProtected As Boolean
    nameIsProtected = False

    Dim i As Integer
    For i = LBound(protectedNames) To UBound(protectedNames)

        If StrComp(sheetName, protectedNames(i), vbTextCompare) = 0 Then
            nameIsProtected = True
            Exit For
        End If

    Next i

    If nameIsProtected Then
        MsgBox "Refused to overwrite protected sheet: " & sheetName, vbExclamation
        Set GetOrCreateWorksheet = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set GetOrCreateWorksheet = ws

End Function

Public Function ParseGeminiDataToSeparateSheets(jsonData As String) As Integer
    On Error GoTo ErrorHandler

    Dim tableCount As Integer
    tableCount = 0

    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    pos = 1

    Call Logging.LogInfo("Starting ParseGeminiDataToSeparateSheets")

    Do While pos < Len(jsonData)
        tableStart = InStr(pos, jsonData, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = Utilities.FindObjectEnd(jsonData, tableStart)
        If tableEnd = 0 Then Exit Do

        Dim tableContent As String
        tableContent = Mid(jsonData, tableStart, tableEnd - tableStart + 1)

        On Error GoTo TableErrorHandler
        tableCount = tableCount + 1
        Call Logging.LogInfo("Parsing table " & tableCount)

        Dim ws As Worksheet
        Set ws = GetOrCreateWorksheet("Table_" & tableCount)

        Call ParseSingleTableToSheet(tableContent, ws)

        ' Reset error handler on success
        On Error GoTo ErrorHandler

        ' Update position to continue
        pos = tableEnd + 1
        GoTo NextTable

TableErrorHandler:
        Call Logging.LogError("? Failed to parse table at position " & tableStart & ": " & Err.Description)
        tableCount = tableCount - 1 ' don’t count failed table
        On Error GoTo ErrorHandler
        pos = tableEnd + 1 ' skip past bad block
        GoTo NextTable

NextTable:
    Loop

    If tableCount = 0 Then
        Set ws = GetOrCreateWorksheet("No_Data")
        ws.Cells(1, 1).Value = "No table data found or all tables failed to parse"
        ws.Cells(1, 1).Font.Italic = True
        tableCount = 1
    End If

    ParseGeminiDataToSeparateSheets = tableCount
    Exit Function

ErrorHandler:
    Call Logging.LogError("Unexpected error parsing Gemini JSON data: " & Err.Description)
    MsgBox "Critical error: " & Err.Description, vbCritical
    ParseGeminiDataToSeparateSheets = 0
End Function

Private Sub ParseSingleTableToSheet(tableJson As String, ws As Worksheet)
    Dim json As Object
    Set json = JsonConverter.ParseJson(tableJson)

    Dim headers As Collection
    Dim rows As Collection
    Set headers = json("headers")
    Set rows = json("rows")

    Dim col As Integer, rowNum As Integer
    rowNum = 1

    ' Write headers
    For col = 1 To headers.Count
        ws.Cells(rowNum, col).Value = headers(col)
        ws.Cells(rowNum, col).Font.Bold = True
    Next col

    ' Write rows
    Dim r As Variant
    For Each r In rows
        rowNum = rowNum + 1
        Dim containsSignature As Boolean
        containsSignature = False

        For col = 1 To headers.Count
            Dim cellVal As Variant
            If col <= r.Count Then
                cellVal = r(col)
                If IsNull(cellVal) Then cellVal = ""
                ws.Cells(rowNum, col).Value = cellVal

                ' Check for signature marker
                If LCase(cellVal) = "signature detected" Then
                    containsSignature = True
                End If
            End If
        Next col

        ' Highlight row if signature is detected
        If containsSignature Then
            ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, headers.Count)).Interior.Color = RGB(198, 239, 206) ' Light green
        End If
    Next r
End Sub