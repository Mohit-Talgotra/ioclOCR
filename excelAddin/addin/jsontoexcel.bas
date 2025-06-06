Option Explicit

' Global JavaScript engine for JSON parsing
Private jsEngine As Object
Private Const MODULE_NAME As String = "JSONToExcelConverter"

' Initialize JavaScript engine
Private Sub InitializeJSEngine()
    If jsEngine Is Nothing Then
        Set jsEngine = CreateObject("InternetExplorer.Application")
        jsEngine.Visible = False
        jsEngine.Navigate "about:blank"
        
        ' Wait for page to load
        Do While jsEngine.Busy Or jsEngine.ReadyState <> 4
            DoEvents
        Loop
        
        Debug.Print "JavaScript engine initialized successfully"
    End If
End Sub

' Clean up JavaScript engine
Private Sub CleanupJSEngine()
    If Not jsEngine Is Nothing Then
        jsEngine.Quit
        Set jsEngine = Nothing
        Debug.Print "JavaScript engine cleaned up"
    End If
End Sub

' Enhanced JSON parsing with error handling
Private Function ParseJSON(jsonString As String, path As String) As Variant
    InitializeJSEngine
    
    ' Escape single quotes in JSON string
    Dim escapedJSON As String
    escapedJSON = Replace(jsonString, "'", "\'")
    escapedJSON = Replace(escapedJSON, vbCrLf, "")
    escapedJSON = Replace(escapedJSON, vbLf, "")
    escapedJSON = Replace(escapedJSON, vbCr, "")
    
    Dim jsCode As String
    jsCode = "try { " & _
             "var obj = JSON.parse('" & escapedJSON & "'); " & _
             "var result = obj" & path & "; " & _
             "if (result === undefined) return 'UNDEFINED'; " & _
             "if (result === null) return 'NULL'; " & _
             "if (typeof result === 'object') return JSON.stringify(result); " & _
             "return result; " & _
             "} catch(e) { return 'ERROR: ' + e.message; }"
    
    Dim result As Variant
    result = jsEngine.Document.parentWindow.execScript(jsCode, "JavaScript")
    
    If Left(CStr(result), 6) = "ERROR:" Then
        Debug.Print "JSON Parse Error: " & result & " for path: " & path
        ParseJSON = ""
    ElseIf CStr(result) = "UNDEFINED" Or CStr(result) = "NULL" Then
        ParseJSON = ""
    Else
        ParseJSON = result
    End If
End Function

' Get array length from JSON
Private Function GetJSONArrayLength(jsonString As String, path As String) As Long
    Dim lengthResult As Variant
    lengthResult = ParseJSON(jsonString, path & ".length")
    
    If IsNumeric(lengthResult) Then
        GetJSONArrayLength = CLng(lengthResult)
    Else
        GetJSONArrayLength = 0
    End If
End Function

' Check if JSON path exists
Private Function JSONPathExists(jsonString As String, path As String) As Boolean
    Dim result As String
    result = CStr(ParseJSON(jsonString, path))
    JSONPathExists = (result <> "" And Left(result, 6) <> "ERROR:")
End Function

' Main conversion function
Public Sub ConvertJSONToExcel(jsonFilePath As String, Optional excelOutputPath As String = "")
    On Error GoTo ErrorHandler
    
    Debug.Print "Starting JSON to Excel conversion: " & jsonFilePath
    
    ' Determine output path
    If excelOutputPath = "" Then
        Dim baseName As String
        baseName = Left(Dir(jsonFilePath), InStrRev(Dir(jsonFilePath), ".") - 1)
        excelOutputPath = Replace(jsonFilePath, Dir(jsonFilePath), baseName & ".xlsx")
    End If
    
    ' Read JSON file
    Dim jsonContent As String
    jsonContent = ReadTextFile(jsonFilePath)
    
    If jsonContent = "" Then
        MsgBox "Error: Could not read JSON file or file is empty", vbCritical
        Exit Sub
    End If
    
    ' Create new workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ' Remove default sheets
    Application.DisplayAlerts = False
    Do While wb.Sheets.Count > 1
        wb.Sheets(wb.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Process pages
    Dim pagesCount As Long
    pagesCount = GetJSONArrayLength(jsonContent, ".pages")
    
    If pagesCount = 0 Then
        MsgBox "No pages found in JSON file", vbWarning
        GoTo Cleanup
    End If
    
    Debug.Print "Processing " & pagesCount & " pages"
    
    Dim i As Long
    For i = 0 To pagesCount - 1
        ProcessPage wb, jsonContent, i
    Next i
    
    ' Remove the default sheet if it still exists
    If wb.Sheets.Count > pagesCount Then
        Application.DisplayAlerts = False
        wb.Sheets(1).Delete
        Application.DisplayAlerts = True
    End If
    
    ' Save workbook
    wb.SaveAs excelOutputPath
    Debug.Print "Excel file saved to: " & excelOutputPath
    
    MsgBox "Conversion completed successfully!" & vbCrLf & "Output: " & excelOutputPath, vbInformation
    
Cleanup:
    CleanupJSEngine
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ConvertJSONToExcel: " & Err.Description
    MsgBox "Error during conversion: " & Err.Description, vbCritical
    CleanupJSEngine
End Sub

' Process individual page
Private Sub ProcessPage(wb As Workbook, jsonContent As String, pageIndex As Long)
    On Error GoTo ErrorHandler
    
    ' Get page data
    Dim pageNumberResult As Variant
    pageNumberResult = ParseJSON(jsonContent, ".pages[" & pageIndex & "].page_number")
    
    Dim pageNumber As Long
    If IsNumeric(pageNumberResult) Then
        pageNumber = CLng(pageNumberResult)
    Else
        pageNumber = pageIndex + 1
    End If
    
    ' Create worksheet
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Page " & pageNumber
    
    Debug.Print "Processing Page " & pageNumber
    
    ' Format page content
    FormatPageWorksheet ws, jsonContent, pageIndex, pageNumber
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error processing page " & pageIndex & ": " & Err.Description
End Sub

' Format worksheet with page content
Private Sub FormatPageWorksheet(ws As Worksheet, jsonContent As String, pageIndex As Long, pageNumber As Long)
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = 1
    
    ' Get page content path
    Dim contentPath As String
    contentPath = ".pages[" & pageIndex & "].content"
    
    ' Document type
    Dim docType As String
    docType = CStr(ParseJSON(jsonContent, contentPath & ".document_type"))
    If docType = "" Then docType = "Unknown Document Type"
    
    ' Add document header
    ws.Cells(currentRow, 1).Value = "Document Type: " & docType
    FormatTitleCell ws.Cells(currentRow, 1)
    currentRow = currentRow + 1
    
    ' Page metadata - header
    Dim pageHeader As String
    pageHeader = CStr(ParseJSON(jsonContent, contentPath & ".page_metadata.header"))
    If pageHeader <> "" Then
        ws.Cells(currentRow, 1).Value = "Header: " & pageHeader
        currentRow = currentRow + 1
    End If
    
    ' Page number
    ws.Cells(currentRow, 1).Value = "Page: " & pageNumber
    currentRow = currentRow + 2
    
    ' Process tables
    currentRow = ProcessTables(ws, jsonContent, contentPath, currentRow)
    
    ' Process sections
    currentRow = ProcessSections(ws, jsonContent, contentPath, currentRow)
    
    ' Process key-value pairs
    currentRow = ProcessKeyValuePairs(ws, jsonContent, contentPath, currentRow)
    
    ' Auto-fit columns
    ws.Columns("A:J").AutoFit
    
    ' Set minimum column widths
    Dim col As Long
    For col = 1 To 10
        If ws.Columns(col).ColumnWidth < 12 Then
            ws.Columns(col).ColumnWidth = 12
        End If
        If ws.Columns(col).ColumnWidth > 50 Then
            ws.Columns(col).ColumnWidth = 50
        End If
    Next col
    
    Debug.Print "Page " & pageNumber & " formatted successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error formatting page " & pageNumber & ": " & Err.Description
End Sub

' Process tables in the page
Private Function ProcessTables(ws As Worksheet, jsonContent As String, contentPath As String, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim tablesCount As Long
    tablesCount = GetJSONArrayLength(jsonContent, contentPath & ".tables")
    
    If tablesCount = 0 Then
        ProcessTables = currentRow
        Exit Function
    End If
    
    Debug.Print "Processing " & tablesCount & " tables"
    
    Dim i As Long
    For i = 0 To tablesCount - 1
        Dim tablePath As String
        tablePath = contentPath & ".tables[" & i & "]"
        
        ' Table title
        Dim tableTitle As String
        tableTitle = CStr(ParseJSON(jsonContent, tablePath & ".table_title"))
        If tableTitle = "" Then tableTitle = "Table " & (i + 1)
        
        ws.Cells(currentRow, 1).Value = tableTitle
        FormatTitleCell ws.Cells(currentRow, 1)
        currentRow = currentRow + 1
        
        ' Process table headers
        Dim headersJSON As String
        headersJSON = CStr(ParseJSON(jsonContent, tablePath & ".headers"))
        
        If headersJSON <> "" And headersJSON <> "[]" Then
            Dim headers As Variant
            headers = ParseJSONArray(headersJSON)
            
            ' Add headers
            Dim col As Long
            For col = 0 To UBound(headers)
                With ws.Cells(currentRow, col + 1)
                    .Value = headers(col)
                    FormatHeaderCell ws.Cells(currentRow, col + 1)
                End With
            Next col
            currentRow = currentRow + 1
            
            ' Process table data
            Dim dataCount As Long
            dataCount = GetJSONArrayLength(jsonContent, tablePath & ".data")
            
            Dim row As Long
            For row = 0 To dataCount - 1
                Dim rowDataJSON As String
                rowDataJSON = CStr(ParseJSON(jsonContent, tablePath & ".data[" & row & "]"))
                
                If rowDataJSON <> "" And rowDataJSON <> "[]" Then
                    Dim rowData As Variant
                    rowData = ParseJSONArray(rowDataJSON)
                    
                    For col = 0 To UBound(rowData)
                        With ws.Cells(currentRow, col + 1)
                            .Value = rowData(col)
                            FormatDataCell ws.Cells(currentRow, col + 1)
                        End With
                    Next col
                    currentRow = currentRow + 1
                End If
            Next row
        End If
        
        currentRow = currentRow + 1 ' Space between tables
    Next i
    
    ProcessTables = currentRow
    Exit Function
    
ErrorHandler:
    Debug.Print "Error processing tables: " & Err.Description
    ProcessTables = currentRow
End Function

' Process sections in the page
Private Function ProcessSections(ws As Worksheet, jsonContent As String, contentPath As String, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim sectionsCount As Long
    sectionsCount = GetJSONArrayLength(jsonContent, contentPath & ".sections")
    
    If sectionsCount = 0 Then
        ProcessSections = currentRow
        Exit Function
    End If
    
    Debug.Print "Processing " & sectionsCount & " sections"
    
    Dim i As Long
    For i = 0 To sectionsCount - 1
        Dim sectionPath As String
        sectionPath = contentPath & ".sections[" & i & "]"
        
        Dim sectionType As String
        sectionType = CStr(ParseJSON(jsonContent, sectionPath & ".section_type"))
        
        ' Skip table sections (handled separately)
        If sectionType = "table" Then GoTo NextSection
        
        Dim sectionTitle As String
        sectionTitle = CStr(ParseJSON(jsonContent, sectionPath & ".section_title"))
        If sectionTitle = "" Then sectionTitle = "Section " & (i + 1)
        
        ws.Cells(currentRow, 1).Value = sectionTitle
        FormatTitleCell ws.Cells(currentRow, 1)
        currentRow = currentRow + 1
        
        Dim sectionContent As String
        sectionContent = CStr(ParseJSON(jsonContent, sectionPath & ".content"))
        
        Select Case sectionType
            Case "text"
                ws.Cells(currentRow, 1).Value = sectionContent
                currentRow = currentRow + 2
                
            Case "form"
                currentRow = ProcessFormContent(ws, sectionContent, currentRow)
                
            Case "chart"
                ws.Cells(currentRow, 1).Value = "Chart: " & sectionContent
                currentRow = currentRow + 2
                
            Case Else
                ws.Cells(currentRow, 1).Value = sectionContent
                currentRow = currentRow + 2
        End Select
        
NextSection:
    Next i
    
    ProcessSections = currentRow
    Exit Function
    
ErrorHandler:
    Debug.Print "Error processing sections: " & Err.Description
    ProcessSections = currentRow
End Function

' Process form content (key-value pairs)
Private Function ProcessFormContent(ws As Worksheet, contentJSON As String, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    If contentJSON = "" Or contentJSON = "{}" Then
        ProcessFormContent = currentRow
        Exit Function
    End If
    
    ' Parse form content as key-value pairs
    Dim keys As Variant
    keys = GetJSONKeys(contentJSON)
    
    If IsArray(keys) Then
        Dim i As Long
        For i = 0 To UBound(keys)
            Dim key As String
            key = CStr(keys(i))
            
            Dim value As String
            value = CStr(ParseJSONValue(contentJSON, key))
            
            ws.Cells(currentRow, 1).Value = key
            ws.Cells(currentRow, 1).Font.Bold = True
            ws.Cells(currentRow, 2).Value = value
            currentRow = currentRow + 1
        Next i
    End If
    
    ProcessFormContent = currentRow + 1
    Exit Function
    
ErrorHandler:
    Debug.Print "Error processing form content: " & Err.Description
    ProcessFormContent = currentRow
End Function

' Process key-value pairs
Private Function ProcessKeyValuePairs(ws As Worksheet, jsonContent As String, contentPath As String, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim kvpJSON As String
    kvpJSON = CStr(ParseJSON(jsonContent, contentPath & ".key_value_pairs"))
    
    If kvpJSON = "" Or kvpJSON = "{}" Then
        ProcessKeyValuePairs = currentRow
        Exit Function
    End If
    
    ws.Cells(currentRow, 1).Value = "Additional Information"
    FormatTitleCell ws.Cells(currentRow, 1)
    currentRow = currentRow + 1
    
    Dim keys As Variant
    keys = GetJSONKeys(kvpJSON)
    
    If IsArray(keys) Then
        Dim i As Long
        For i = 0 To UBound(keys)
            Dim key As String
            key = CStr(keys(i))
            
            Dim value As String
            value = CStr(ParseJSONValue(kvpJSON, key))
            
            ws.Cells(currentRow, 1).Value = key
            ws.Cells(currentRow, 1).Font.Bold = True
            ws.Cells(currentRow, 2).Value = value
            currentRow = currentRow + 1
        Next i
    End If
    
    ProcessKeyValuePairs = currentRow
    Exit Function
    
ErrorHandler:
    Debug.Print "Error processing key-value pairs: " & Err.Description
    ProcessKeyValuePairs = currentRow
End Function

' Helper function to parse JSON arrays
Private Function ParseJSONArray(jsonArrayString As String) As Variant
    InitializeJSEngine
    
    Dim jsCode As String
    jsCode = "try { " & _
             "var arr = " & jsonArrayString & "; " & _
             "arr.join('|||'); " & _
             "} catch(e) { 'ERROR: ' + e.message; }"
    
    Dim result As String
    result = jsEngine.Document.parentWindow.execScript(jsCode, "JavaScript")
    
    If Left(result, 6) = "ERROR:" Then
        ParseJSONArray = Array()
    Else
        ParseJSONArray = Split(result, "|||")
    End If
End Function

' Helper function to get JSON object keys
Private Function GetJSONKeys(jsonString As String) As Variant
    InitializeJSEngine
    
    Dim jsCode As String
    jsCode = "try { " & _
             "var obj = " & jsonString & "; " & _
             "Object.keys(obj).join('|||'); " & _
             "} catch(e) { 'ERROR: ' + e.message; }"
    
    Dim result As String
    result = jsEngine.Document.parentWindow.execScript(jsCode, "JavaScript")
    
    If Left(result, 6) = "ERROR:" Or result = "" Then
        GetJSONKeys = Array()
    Else
        GetJSONKeys = Split(result, "|||")
    End If
End Function

' Helper function to parse JSON value by key
Private Function ParseJSONValue(jsonString As String, key As String) As Variant
    InitializeJSEngine
    
    Dim jsCode As String
    jsCode = "try { " & _
             "var obj = " & jsonString & "; " & _
             "var result = obj['" & key & "']; " & _
             "typeof result === 'object' ? JSON.stringify(result) : result; " & _
             "} catch(e) { 'ERROR: ' + e.message; }"
    
    ParseJSONValue = jsEngine.Document.parentWindow.execScript(jsCode, "JavaScript")
End Function

' Cell formatting functions
Private Sub FormatTitleCell(cell As Range)
    With cell
        .Font.Bold = True
        .Font.Size = 14
        .Font.Name = "Calibri"
    End With
End Sub

Private Sub FormatHeaderCell(cell As Range)
    With cell
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .Interior.Color = RGB(221, 235, 247) ' Light blue
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub FormatDataCell(cell As Range)
    With cell
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Name = "Calibri"
    End With
End Sub

' Read text file function
Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ReadTextFile = fileContent
    Exit Function
    
ErrorHandler:
    Debug.Print "Error reading file " & filePath & ": " & Err.Description
    If fileNum > 0 Then Close #fileNum
    ReadTextFile = ""
End Function

' Public interface function
Public Sub ConvertJSONFileToExcel()
    Dim filePath As String
    filePath = Application.GetOpenFilename("JSON Files (*.json), *.json", , "Select JSON File to Convert")
    
    If filePath <> "False" Then
        ConvertJSONToExcel filePath
    End If
End Sub