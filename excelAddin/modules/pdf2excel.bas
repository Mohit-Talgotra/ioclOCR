Option Explicit

Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub PDFToExcel(control As IRibbonControl)

    Dim pdfPath As String
    Dim pdfFolder As String
    Dim imageFolder As String
    Dim shellCmd As String
    Dim wsh As Object
    Dim exitCode As Long

    pdfPath = GetPDFFile()
    If pdfPath = "" Then
        Debug.Print "No PDF file selected."
        Exit Sub
    End If

    pdfFolder = Left(pdfPath, InStrRev(pdfPath, "\") - 1)
    imageFolder = "C:\Users\talgo\AppData\Roaming\Microsoft\AddIns\pdf_images"

    Debug.Print "PDF path: " & pdfPath
    Debug.Print "PDF folder: " & pdfFolder
    Debug.Print "Image folder: " & imageFolder

    If Dir(imageFolder, vbDirectory) = "" Then
        MkDir imageFolder
        Debug.Print "Created image folder: " & imageFolder
    Else
        Debug.Print "Image folder already exists."
    End If

    shellCmd = "cmd /c cd /d """ & imageFolder & """ && pdftoppm -jpeg """ & pdfPath & """ page"
    Debug.Print "Shell command: " & shellCmd

    Set wsh = CreateObject("WScript.Shell")

    exitCode = wsh.Run(shellCmd, 0, True)

    Debug.Print "pdftoppm exited with code: " & exitCode

    If exitCode = 0 Then
        ' Do nothing
    Else
        MsgBox "Error converting PDF to images. Exit code: " & exitCode, vbCritical
    End If

    Set wsh = Nothing
    
    ConvertPDFToExcel (imageFolder)

End Sub

Private Function GetPDFFile() As String
    Dim fileDialog As fileDialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Select PDF File to Extract Tables"
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

Private Sub ConvertPDFToExcel(imageFolder As String)
    Dim tableData As String
    Dim apiKey As String

    apiKey = "AIzaSyA78zCKSrZHcRLB1nvhPJuDPqFeHT8Iu4Q"

    imageFolder = "C:\Users\talgo\AppData\Roaming\Microsoft\AddIns\pdf_images"
    
    If Right(imageFolder, 1) <> "\" Then
        imageFolder = imageFolder & "\"
    End If

    Application.StatusBar = "Extracting tables from images..."
    Application.ScreenUpdating = False

    tableData = ExtractTablesWithGeminiFromImages(imageFolder, apiKey)
    Dim filePath As String
    Dim fileNumber As Integer
    
    filePath = "C:\Users\talgo\AppData\Roaming\Microsoft\AddIns\output\output.txt"
    
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, tableData
    Close #fileNumber
    
    TestParseFromFile (filePath)

    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Function ExtractTablesWithGeminiFromImages(imageFolder As String, apiKey As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim combinedData As String
    Dim pageNum As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Dir(imageFolder, vbDirectory) = "" Then
        MsgBox "Selected folder does not exist or is inaccessible.", vbCritical
        Exit Function
    End If
    
    Set folder = fso.GetFolder(imageFolder)
    If folder Is Nothing Then
        MsgBox "Failed to access folder.", vbCritical
        Exit Function
    End If

    pageNum = 1

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Then
            Debug.Print "Processing page image: " & file.Name
            Dim pageResponse As String
            pageResponse = ProcessImagePageWithGemini(file.Path, apiKey)
            If pageResponse <> "" Then
                combinedData = combinedData & pageResponse & vbCrLf
            End If
            pageNum = pageNum + 1
        End If
    Next file

    ExtractTablesWithGeminiFromImages = combinedData
End Function

Private Function ProcessImagePageWithGemini(imagePath As String, apiKey As String) As String
    Dim fileUri As String
    Dim jsonRequest As String
    Dim responseText As String
    Dim http As Object

    fileUri = UploadFileToGemini(imagePath, apiKey)
    If fileUri = "" Then Exit Function

    jsonRequest = CreateGeminiImageRequest(fileUri)
    
    LogRequestToFile jsonRequest

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 60000, 60000, 60000, 300000
    http.Open "POST", "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" & apiKey, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send jsonRequest

    responseText = http.responseText
    Debug.Print "Gemini Image Response: " & responseText

    If http.Status = 200 Then
        ProcessImagePageWithGemini = ParseGeminiResponse(responseText)
    Else
        Debug.Print "Gemini API Error: " & http.Status & " - " & http.StatusText
        ProcessImagePageWithGemini = ""
    End If

    Set http = Nothing
End Function

Private Function UploadFileToGemini(filePath As String, apiKey As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim boundary As String
    Dim fileStream As Object
    Dim requestStream As Object
    Dim uploadUrl As String
    Dim formHeader1() As Byte
    Dim formHeader2() As Byte
    Dim formFooter() As Byte
    Dim buffer() As Byte
    Dim fileData() As Byte
    Dim metadata As String
    Dim response As String
    Dim uriStart As Long, uriEnd As Long
    Dim fileName As String
    Dim contentType As String

    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    If LCase(Right(fileName, 4)) = ".jpg" Or LCase(Right(fileName, 5)) = ".jpeg" Then
        contentType = "image/jpeg"
    ElseIf LCase(Right(fileName, 4)) = ".png" Then
        contentType = "image/png"
    ElseIf LCase(Right(fileName, 4)) = ".pdf" Then
        contentType = "application/pdf"
    Else
        contentType = "application/octet-stream"
    End If

    boundary = "FormBoundary" & Format(Now, "yyyymmddhhmmss")

    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1
    fileStream.Open
    fileStream.LoadFromFile filePath
    fileData = fileStream.Read

    Dim formData As String
    formData = "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
               "Content-Type: " & contentType & vbCrLf & vbCrLf

    Dim formDataBytes() As Byte
    Dim footerBytes() As Byte
    formDataBytes = StrConv(formData, vbFromUnicode)
    footerBytes = StrConv(vbCrLf & "--" & boundary & "--" & vbCrLf, vbFromUnicode)

    Set requestStream = CreateObject("ADODB.Stream")
    requestStream.Type = 1
    requestStream.Open

    requestStream.Write formDataBytes
    requestStream.Write fileData
    requestStream.Write footerBytes
    requestStream.Position = 0

    buffer = requestStream.Read(requestStream.Size)

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 60000, 60000, 60000, 300000
    uploadUrl = "https://generativelanguage.googleapis.com/upload/v1beta/files?key=" & apiKey
    http.Open "POST", uploadUrl, False
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send buffer

    Debug.Print "Upload Status: " & http.Status
    Debug.Print "Upload Response: " & http.responseText
    
    If http.Status = 200 Then
        response = http.responseText
        uriStart = InStr(response, """uri"": """)
        If uriStart = 0 Then uriStart = InStr(response, """uri"":""")
        
        If uriStart > 0 Then
            uriStart = uriStart + Len("""uri"": """)
            If uriStart = InStr(response, """uri"":""") + Len("""uri"":""") Then
                uriStart = InStr(response, """uri"":""") + Len("""uri"":""")
            End If
            uriEnd = InStr(uriStart, response, """")
            If uriEnd > uriStart Then
                UploadFileToGemini = Mid(response, uriStart, uriEnd - uriStart)
                Debug.Print "Extracted URI: " & UploadFileToGemini
                GoTo Cleanup
            End If
        End If
        
        UploadFileToGemini = ""
    Else
        MsgBox "File Upload Error: " & http.Status & " - " & http.StatusText & vbCrLf & http.responseText, vbCritical
        UploadFileToGemini = ""
    End If

Cleanup:
    fileStream.Close
    requestStream.Close
    Set fileStream = Nothing
    Set requestStream = Nothing
    Set http = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Unexpected error during upload: " & Err.Description, vbCritical
    UploadFileToGemini = ""
    On Error Resume Next
    If Not fileStream Is Nothing Then fileStream.Close
    If Not requestStream Is Nothing Then requestStream.Close
    Set fileStream = Nothing
    Set requestStream = Nothing
    Set http = Nothing
End Function

Private Function CreateGeminiImageRequest(fileUri As String) As String
   Dim jsonRequest As String
   Dim promptText As String

   promptText = "Extract all tables from this image. Return ONLY a JSON array in this exact format: "
   promptText = promptText & "[{" & Chr(34) & "headers" & Chr(34) & ": [" & Chr(34) & "header1" & Chr(34) & ", " & Chr(34) & "header2" & Chr(34) & ", ...], "
   promptText = promptText & Chr(34) & "rows" & Chr(34) & ": [[" & Chr(34) & "value1" & Chr(34) & ", " & Chr(34) & "value2" & Chr(34) & ", ...], "
   promptText = promptText & "[" & Chr(34) & "value1" & Chr(34) & ", " & Chr(34) & "value2" & Chr(34) & ", ...], ...]}]. "
   promptText = promptText & "Do not include any other text, markdown formatting, or code blocks. Each table should be a separate object in the array."
   promptText = promptText & "VERY IMPORTANT: If a signature is detected in a column like User Sign or anything of that kind, put the text SIGNATURE DETECTED in that column in each row where the signature is detected in your structure output"
   promptText = promptText & "Before sending final output, understand the data in context and recheck the output for any misalignment of columns or misplaced data, and fix those problems."

   promptText = Replace(promptText, "\", "\\")
   promptText = Replace(promptText, """", "\""")
   promptText = Replace(promptText, vbCrLf, "\n")
   promptText = Replace(promptText, vbLf, "\n")
   promptText = Replace(promptText, vbCr, "\n")
   
   jsonRequest = "{" & _
       """contents"": [" & _
           "{" & _
               """role"": ""user""," & _
               """parts"": [" & _
                   "{" & _
                       """text"": """ & promptText & """" & _
                   "}," & _
                   "{" & _
                       """file_data"": {" & _
                           """mime_type"": ""image/jpeg""," & _
                           """file_uri"": """ & fileUri & """" & _
                       "}" & _
                   "}" & _
               "]" & _
           "}" & _
       "]" & _
   "}"
   
   CreateGeminiImageRequest = jsonRequest
End Function

Private Sub LogRequestToFile(jsonRequest As String)
   On Error Resume Next
   
   Dim filePath As String
   Dim fileNumber As Integer
   Dim promptStart As Long, promptEnd As Long
   Dim extractedPrompt As String
   
   filePath = "C:\Users\talgo\OneDrive\Desktop\request_log.txt"

   promptStart = InStr(jsonRequest, """text"": """) + 9
   promptEnd = InStr(promptStart, jsonRequest, """},")
   
   If promptEnd > promptStart Then
       extractedPrompt = Mid(jsonRequest, promptStart, promptEnd - promptStart - 1)
       extractedPrompt = Replace(extractedPrompt, "\""", """")
       extractedPrompt = Replace(extractedPrompt, "\\", "\")
       extractedPrompt = Replace(extractedPrompt, "\n", vbCrLf)
   Else
       extractedPrompt = "Could not extract prompt text"
   End If
   
   fileNumber = FreeFile
   Open filePath For Output As #fileNumber
   Print #fileNumber, "=== PROMPT TEXT SENT TO GEMINI ==="
   Print #fileNumber, extractedPrompt
   Print #fileNumber, ""
   Print #fileNumber, "=== FULL JSON REQUEST ==="
   Print #fileNumber, jsonRequest
   Close #fileNumber
End Sub

Private Function ParseGeminiResponse(response As String) As String
    On Error GoTo ErrorHandler

    Dim textStart As Long, textEnd As Long
    Dim searchPattern As String
    searchPattern = """text"": """
    
    textStart = InStr(response, searchPattern)
    
    If textStart > 0 Then
        textStart = textStart + Len(searchPattern)

        Dim pos As Long, inEscape As Boolean
        pos = textStart
        inEscape = False
        
        Do While pos <= Len(response)
            Dim currentChar As String
            currentChar = Mid(response, pos, 1)
            
            If currentChar = "\" And Not inEscape Then
                inEscape = True
            ElseIf currentChar = """" And Not inEscape Then
                If pos + 1 <= Len(response) Then
                    Dim nextChars As String
                    nextChars = Mid(response, pos + 1, 10)
                    If InStr(nextChars, "}") > 0 Or InStr(nextChars, "]") > 0 Then
                        textEnd = pos
                        Exit Do
                    End If
                End If
            Else
                inEscape = False
            End If
            pos = pos + 1
        Loop
        
        If textEnd > textStart Then
            Dim extractedText As String
            extractedText = Mid(response, textStart, textEnd - textStart)
            
            extractedText = Replace(extractedText, "\n", " ")
            extractedText = Replace(extractedText, "\r", " ")
            extractedText = Replace(extractedText, "\t", " ")
            extractedText = Replace(extractedText, "\""", """")
            extractedText = Replace(extractedText, "\\", "\")
            
            Dim jsonStart As Long, jsonEnd As Long
            jsonStart = InStr(extractedText, "[")
            jsonEnd = InStrRev(extractedText, "]")
            
            If jsonStart > 0 And jsonEnd > jsonStart Then
                ParseGeminiResponse = Mid(extractedText, jsonStart, jsonEnd - jsonStart + 1)
            Else
                ParseGeminiResponse = extractedText
            End If
        End If
    Else
        ParseGeminiResponse = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ParseGeminiResponse = ""
End Function

Private Sub TestParseFromFile(filePath As String)
    Dim fileContent As String
    Dim fileNumber As Integer

    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContent = Input(LOF(fileNumber), fileNumber)
    Close #fileNumber

    If fileContent <> "" Then
        Dim tableCount As Integer
        tableCount = ParseGeminiDataToSeparateSheets(fileContent)
        MsgBox "Test parsing completed successfully! Created " & tableCount & " sheets for " & tableCount & " tables.", vbInformation
    Else
        MsgBox "File is empty or could not be read.", vbCritical
    End If
End Sub

Private Sub ParseGeminiTableJSON(jsonString As String, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long
    currentRow = 1
    
    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    
    pos = 1
    
    Do While pos < Len(jsonString)
        tableStart = InStr(pos, jsonString, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = FindObjectEnd(jsonString, tableStart)
        If tableEnd = 0 Then Exit Do
        
        Dim tableContent As String
        tableContent = Mid(jsonString, tableStart, tableEnd - tableStart + 1)

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
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error populating data: " & Err.Description, vbCritical
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
    Exit Sub
    
ErrorHandler:
    ReDim items(0 To 0)
    items(0) = ""
    ParseJSONStringArray = items
End Sub

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

Private Function FindObjectEnd(text As String, startPos As Long) As Long
    Dim braceCount As Long
    Dim i As Long
    braceCount = 1
    
    For i = startPos + 1 To Len(text)
        If Mid(text, i, 1) = "{" Then
            braceCount = braceCount + 1
        ElseIf Mid(text, i, 1) = "}" Then
            braceCount = braceCount - 1
            If braceCount = 0 Then
                FindObjectEnd = i
                Exit Function
            End If
        End If
    Next i
    
    FindObjectEnd = 0
End Function

Public Sub ParseGeminiTableJSONToSheet(jsonString As String, ws As Worksheet)
    On Error GoTo ErrorHandler

    ws.Cells.Clear
    
    Dim currentRow As Long
    currentRow = 1

    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    
    pos = 1

    Do While pos < Len(jsonString)
        tableStart = InStr(pos, jsonString, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = FindObjectEnd(jsonString, tableStart)
        If tableEnd = 0 Then Exit Do
        
        Dim tableContent As String
        tableContent = Mid(jsonString, tableStart, tableEnd - tableStart + 1)

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
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing Gemini table JSON to sheet: " & Err.Description, vbCritical
End Sub

Public Sub TestParseMultipleFiles()
    Dim basePath As String
    Dim pageNum As Integer
    Dim filePath As String
    Dim ws As Worksheet
    
    basePath = "C:\Users\talgo\OneDrive\Desktop\output_page_"
    pageNum = 1

    Do
        filePath = basePath & pageNum & ".txt"
        
        If Dir(filePath) = "" Then Exit Do

        Set ws = GetOrCreateWorksheet("Page_" & pageNum)

        Dim fileContent As String
        Dim fileNumber As Integer
        
        fileNumber = FreeFile
        Open filePath For Input As #fileNumber
        fileContent = Input(LOF(fileNumber), fileNumber)
        Close #fileNumber
        
        If fileContent <> "" Then
            ParseGeminiTableJSONToSheet fileContent, ws
        End If
        
        pageNum = pageNum + 1
    Loop
    
    If pageNum > 1 Then
        MsgBox "Parsed " & (pageNum - 1) & " files into separate sheets!", vbInformation
    Else
        MsgBox "No output files found. Expected files like: " & basePath & "1.txt", vbCritical
    End If
End Sub

Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    
    Set GetOrCreateWorksheet = ws
End Function

Private Function ParseGeminiDataToSeparateSheets(jsonData As String) As Integer
    On Error GoTo ErrorHandler
    
    Dim tableCount As Integer
    tableCount = 0
    
    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    
    pos = 1
    
    Do While pos < Len(jsonData)
        tableStart = InStr(pos, jsonData, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = FindObjectEnd(jsonData, tableStart)
        If tableEnd = 0 Then Exit Do
        
        Dim tableContent As String
        tableContent = Mid(jsonData, tableStart, tableEnd - tableStart + 1)

        tableCount = tableCount + 1
        Dim ws As Worksheet
        Set ws = GetOrCreateWorksheet("Table_" & tableCount)

        ParseSingleTableToSheet tableContent, ws

        pos = tableEnd + 1
    Loop

    If tableCount = 0 Then
        Set ws = GetOrCreateWorksheet("No_Data")
        ws.Cells(1, 1).Value = "No table data found in the response"
        ws.Cells(1, 1).Font.Italic = True
        tableCount = 1
    End If
    
    ParseGeminiDataToSeparateSheets = tableCount
    Exit Function
    
ErrorHandler:
    MsgBox "Error parsing Gemini data to separate sheets: " & Err.Description, vbCritical
    ParseGeminiDataToSeparateSheets = 0
End Function

Private Sub ParseSingleTableToSheet(tableContent As String, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ws.Cells.Clear
    
    Dim currentRow As Long
    currentRow = 1
    
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
    
    If currentRow = 1 Then
        ws.Cells(currentRow, 1).Value = "No table data found in this table object"
        ws.Cells(currentRow, 1).Font.Italic = True
    End If
    
    ws.Columns.AutoFit
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing single table to sheet: " & Err.Description, vbCritical
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