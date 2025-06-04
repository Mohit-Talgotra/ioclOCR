Option Explicit

Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub ConvertPDFToExcel(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Dim pdfFilePath As String
    pdfFilePath = GetPDFFile()

    If pdfFilePath = "" Then
        MsgBox "No file selected.", vbInformation
        Exit Sub
    End If

    Application.StatusBar = "Converting PDF to Excel..."
    Application.ScreenUpdating = False

    Dim excelData As String
    excelData = UploadPDFToFlask(pdfFilePath)

    If excelData <> "" Then
        PopulateExcelWithData excelData
        MsgBox "PDF conversion completed successfully!", vbInformation
    Else
        MsgBox "Failed to convert PDF. Please check your Flask server.", vbCritical
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function GetPDFFile() As String
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Select PDF File to Convert"
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        
        If .Show = -1 Then
            GetPDFFile = .SelectedItems(1)
        Else
            GetPDFFile = ""
        End If
    End With
    
    Set fileDialog = Nothing
End Function

Private Function UploadPDFToFlask(filePath As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim stream As Object
    Dim boundary As String
    Dim CRLF As String
    Dim postData() As Byte
    Dim fileData() As Byte
    Dim headerBytes() As Byte
    Dim trailerBytes() As Byte
    Dim fileName As String
    Dim header As String
    Dim trailer As String
    Dim totalSize As Long
    Dim pos As Long
    Dim i As Long

    boundary = "----WebKitFormBoundary" & Format(Now, "yyyymmddhhmmss")
    CRLF = vbCrLf
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath
    fileData = stream.Read
    stream.Close
    Set stream = Nothing

    header = "--" & boundary & CRLF & _
             "Content-Disposition: form-data; name=" & Chr(34) & "file" & Chr(34) & "; filename=" & Chr(34) & fileName & Chr(34) & CRLF & _
             "Content-Type: application/pdf" & CRLF & CRLF

    trailer = CRLF & "--" & boundary & "--" & CRLF

    headerBytes = StrConv(header, vbFromUnicode)
    trailerBytes = StrConv(trailer, vbFromUnicode)

    totalSize = UBound(headerBytes) + 1 + UBound(fileData) + 1 + UBound(trailerBytes) + 1

    ReDim postData(0 To totalSize - 1)
    pos = 0

    For i = 0 To UBound(headerBytes)
        postData(pos) = headerBytes(i)
        pos = pos + 1
    Next i

    For i = 0 To UBound(fileData)
        postData(pos) = fileData(i)
        pos = pos + 1
    Next i

    For i = 0 To UBound(trailerBytes)
        postData(pos) = trailerBytes(i)
        pos = pos + 1
    Next i

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 60000, 60000, 60000, 300000
    http.Option(4) = &H3300
    http.Open "POST", "https://127.0.0.1:5000/convert-pdf", False  ' Changed to http
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.SetRequestHeader "Content-Length", CStr(totalSize)
    
    http.Send postData

    If http.Status = 200 Then
        UploadPDFToFlask = http.ResponseText
    Else
        MsgBox "Server Error: " & http.Status & " - " & http.StatusText & vbCrLf & http.ResponseText, vbCritical
        UploadPDFToFlask = ""
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    If Not http Is Nothing Then Set http = Nothing
    If Not stream Is Nothing Then Set stream = Nothing
    MsgBox "Error uploading file: " & Err.Description, vbCritical
    UploadPDFToFlask = ""
End Function

Private Function ReadFileAsBase64(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim base64String As String
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    
    ReDim fileData(LOF(fileNum) - 1)
    Get #fileNum, , fileData
    Close #fileNum
    
    base64String = EncodeBase64(fileData)
    ReadFileAsBase64 = base64String
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ReadFileAsBase64 = ""
End Function

Private Function EncodeBase64(data() As Byte) As String
    On Error GoTo ErrorHandler
    
    Dim xml As Object
    Dim node As Object
    
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("base64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = data
    EncodeBase64 = node.text
    
    Set node = Nothing
    Set xml = Nothing
    Exit Function
    
ErrorHandler:
    EncodeBase64 = ""
End Function

Private Sub PopulateExcelWithData(jsonData As String)
    On Error GoTo ErrorHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Clear existing data in current sheet?", vbYesNoCancel + vbQuestion)
    
    If response = vbCancel Then Exit Sub
    If response = vbYes Then ActiveSheet.Cells.Clear
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If InStr(jsonData, """data""") > 0 Then        
        ParseAndPopulateJSON jsonData, ws
    Else
        ParseAndPopulateCSV jsonData, ws
    End If
    
    ws.Columns.AutoFit
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error populating data: " & Err.Description, vbCritical
End Sub

Private Sub ParseAndPopulateJSON(jsonString As String, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = 1
    
    If InStr(jsonString, """document_metadata""") > 0 Then
        ws.Cells(currentRow, 1).value = "DOCUMENT METADATA"
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Interior.Color = RGB(200, 200, 200)
        currentRow = currentRow + 1
        
        Dim totalPages As String
        totalPages = ExtractJSONValue(jsonString, "total_pages")
        If totalPages <> "" Then
            ws.Cells(currentRow, 1).value = "Total Pages:"
            ws.Cells(currentRow, 2).value = totalPages
            currentRow = currentRow + 2
        End If
    End If
    
    Dim pageStart As Long, pageEnd As Long
    Dim pageContent As String
    Dim pageNum As Long
    
    pageStart = InStr(jsonString, """pages"":[")
    If pageStart > 0 Then
        Dim searchPos As Long
        searchPos = pageStart
        pageNum = 1
        
        Do While InStr(searchPos, jsonString, """page_number"":" & pageNum) > 0
            currentRow = ProcessPage(jsonString, pageNum, ws, currentRow)
            pageNum = pageNum + 1
            searchPos = InStr(searchPos + 1, jsonString, """page_number"":" & pageNum)
            If searchPos = 0 Then Exit Do
        Loop
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing JSON: " & Err.Description, vbCritical
End Sub

Private Function ProcessPage(jsonString As String, pageNum As Long, ws As Worksheet, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    ws.Cells(currentRow, 1).value = "PAGE " & pageNum
    ws.Cells(currentRow, 1).Font.Bold = True
    ws.Cells(currentRow, 1).Interior.Color = RGB(173, 216, 230)
    ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 10)).Merge
    currentRow = currentRow + 1
    
    Dim docType As String
    docType = ExtractPageValue(jsonString, pageNum, "document_type")
    If docType <> "" Then
        ws.Cells(currentRow, 1).value = "Document Type:"
        ws.Cells(currentRow, 2).value = docType
        ws.Cells(currentRow, 2).Font.Bold = True
        currentRow = currentRow + 1
    End If
    
    Dim headerText As String, footerText As String
    headerText = ExtractNestedValue(jsonString, pageNum, "page_metadata", "header")
    footerText = ExtractNestedValue(jsonString, pageNum, "page_metadata", "footer")
    
    If headerText <> "" Then
        ws.Cells(currentRow, 1).value = "Header:"
        ws.Cells(currentRow, 2).value = Replace(headerText, "\n", " | ")
        currentRow = currentRow + 1
    End If
    
    If footerText <> "" Then
        ws.Cells(currentRow, 1).value = "Footer:"
        ws.Cells(currentRow, 2).value = Replace(footerText, "\n", " | ")
        currentRow = currentRow + 1
    End If
    
    currentRow = ProcessTables(jsonString, pageNum, ws, currentRow)
    
    ProcessPage = currentRow + 1
    Exit Function
    
ErrorHandler:
    ProcessPage = currentRow
End Function

Private Function ProcessTables(jsonString As String, pageNum As Long, ws As Worksheet, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow + 1
    
    Dim tablesStart As Long, tablesEnd As Long
    Dim tableContent As String
    
    Dim pageStart As Long
    pageStart = InStr(jsonString, """page_number"":" & pageNum)
    
    If pageStart > 0 Then
        tablesStart = InStr(pageStart, jsonString, """tables"":[")
        If tablesStart > 0 Then
            tablesEnd = FindArrayEnd(jsonString, tablesStart + 9)
            
            If tablesEnd > tablesStart Then
                tableContent = Mid(jsonString, tablesStart + 9, tablesEnd - tablesStart - 9)
                currentRow = ParseTableData(tableContent, ws, currentRow)
            End If
        End If
    End If
    
    ProcessTables = currentRow
    Exit Function
    
ErrorHandler:
    ProcessTables = startRow
End Function

Private Function ParseTableData(tableContent As String, ws As Worksheet, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim tableTitle As String
    tableTitle = ExtractJSONValue(tableContent, "table_title")
    
    If tableTitle <> "" Then
        ws.Cells(currentRow, 1).value = tableTitle
        ws.Cells(currentRow, 1).Font.Bold = True
        ws.Cells(currentRow, 1).Interior.Color = RGB(255, 255, 0)
        currentRow = currentRow + 1
    End If
    
    Dim headersStart As Long, headersEnd As Long
    headersStart = InStr(tableContent, """headers"":[")
    
    If headersStart > 0 Then
        headersEnd = FindArrayEnd(tableContent, headersStart + 11)
        Dim headersContent As String
        headersContent = Mid(tableContent, headersStart + 11, headersEnd - headersStart - 11)
        
        Dim headers As Variant
        headers = ParseJSONArray(headersContent)
        
        Dim col As Long
        For col = 0 To UBound(headers)
            ws.Cells(currentRow, col + 1).value = headers(col)
            ws.Cells(currentRow, col + 1).Font.Bold = True
            ws.Cells(currentRow, col + 1).Interior.Color = RGB(220, 220, 220)
        Next col
        currentRow = currentRow + 1
        
        Dim dataStart As Long, dataEnd As Long
        dataStart = InStr(tableContent, """data"":[")
        
        If dataStart > 0 Then
            dataEnd = FindArrayEnd(tableContent, dataStart + 8)
            Dim dataContent As String
            dataContent = Mid(tableContent, dataStart + 8, dataEnd - dataStart - 8)
            
            currentRow = ParseDataRows(dataContent, ws, currentRow)
        End If
    End If
    
    ParseTableData = currentRow + 1
    Exit Function
    
ErrorHandler:
    ParseTableData = startRow
End Function

Private Function ParseDataRows(dataContent As String, ws As Worksheet, startRow As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim rows As Variant
    Dim rowContent As String
    Dim i As Long, j As Long
    
    Dim pos As Long, nextPos As Long
    pos = 1
    
    Do While pos < Len(dataContent)
        pos = InStr(pos, dataContent, "[")
        If pos = 0 Then Exit Do
        
        nextPos = InStr(pos + 1, dataContent, "]")
        If nextPos = 0 Then Exit Do
        
        rowContent = Mid(dataContent, pos + 1, nextPos - pos - 1)
        
        Dim rowData As Variant
        rowData = ParseJSONArray(rowContent)
        
        For j = 0 To UBound(rowData)
            If rowData(j) <> "null" And rowData(j) <> "" Then
                ws.Cells(currentRow, j + 1).value = rowData(j)
            End If
        Next j
        
        currentRow = currentRow + 1
        pos = nextPos + 1
    Loop
    
    ParseDataRows = currentRow
    Exit Function
    
ErrorHandler:
    ParseDataRows = startRow
End Function

Private Function ExtractJSONValue(jsonText As String, key As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(jsonText, """" & key & """:""")
    If startPos > 0 Then
        startPos = startPos + Len(key) + 4
        endPos = InStr(startPos, jsonText, """")
        If endPos > startPos Then
            ExtractJSONValue = Mid(jsonText, startPos, endPos - startPos)
        End If
    Else
        startPos = InStr(jsonText, """" & key & """:")
        If startPos > 0 Then
            startPos = startPos + Len(key) + 3
            endPos = InStr(startPos, jsonText, ",")
            If endPos = 0 Then endPos = InStr(startPos, jsonText, "}")
            If endPos > startPos Then
                ExtractJSONValue = Trim(Mid(jsonText, startPos, endPos - startPos))
            End If
        End If
    End If
End Function

Private Function ExtractPageValue(jsonText As String, pageNum As Long, key As String) As String
    Dim pageStart As Long
    pageStart = InStr(jsonText, """page_number"":" & pageNum)
    If pageStart > 0 Then
        Dim pageEnd As Long
        pageEnd = InStr(pageStart, jsonText, """page_number"":" & (pageNum + 1))
        If pageEnd = 0 Then pageEnd = Len(jsonText)
        
        Dim pageContent As String
        pageContent = Mid(jsonText, pageStart, pageEnd - pageStart)
        ExtractPageValue = ExtractJSONValue(pageContent, key)
    End If
End Function

Private Function ExtractNestedValue(jsonText As String, pageNum As Long, parentKey As String, childKey As String) As String
    Dim pageStart As Long
    pageStart = InStr(jsonText, """page_number"":" & pageNum)
    If pageStart > 0 Then
        Dim parentStart As Long
        parentStart = InStr(pageStart, jsonText, """" & parentKey & """:{")
        If parentStart > 0 Then
            Dim parentEnd As Long
            parentEnd = InStr(parentStart, jsonText, "}")
            If parentEnd > parentStart Then
                Dim parentContent As String
                parentContent = Mid(jsonText, parentStart, parentEnd - parentStart)
                ExtractNestedValue = ExtractJSONValue(parentContent, childKey)
            End If
        End If
    End If
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

Private Function ParseJSONArray(arrayContent As String) As Variant
    Dim items() As String
    Dim itemCount As Long
    Dim i As Long, pos As Long, nextPos As Long
    Dim inQuotes As Boolean
    Dim currentItem As String
    
    ReDim items(0 To 50)
    itemCount = 0
    pos = 1
    currentItem = ""
    inQuotes = False
    
    For i = 1 To Len(arrayContent)
        Dim char As String
        char = Mid(arrayContent, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            items(itemCount) = Trim(Replace(currentItem, """", ""))
            itemCount = itemCount + 1
            currentItem = ""
        Else
            currentItem = currentItem & char
        End If
    Next i

    If currentItem <> "" Then
        items(itemCount) = Trim(Replace(currentItem, """", ""))
        itemCount = itemCount + 1
    End If

    ReDim Preserve items(0 To itemCount - 1)
    ParseJSONArray = items
End Function

Private Sub ParseAndPopulateCSV(csvData As String, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim rows As Variant
    Dim cols As Variant
    Dim i As Long, j As Long

    rows = Split(csvData, vbCrLf)
    
    For i = 0 To UBound(rows)
        If Trim(rows(i)) <> "" Then
            cols = Split(rows(i), ",")
            For j = 0 To UBound(cols)
                ws.Cells(i + 1, j + 1).value = Trim(cols(j))
            Next j
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing CSV: " & Err.Description, vbCritical
End Sub