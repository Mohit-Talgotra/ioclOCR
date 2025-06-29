Private BASE_RUN_FOLDER As String
Private IMAGE_FOLDER As String
Private OUTPUT_FOLDER As String
Private LOG_FILE_PATH As String

Public Enum LogLevel
    Info = 1
    Warning = 2
    Error = 3
    Debugging = 4
End Enum

Option Explicit

Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Private Sub InitRunFolders()

    Dim basePath As String
    Dim runIndex As Integer
    Dim runFolder As String

    basePath = "C:\IOCL_OCR\Run_"
    runIndex = 1

    Do
        runFolder = basePath & Format(runIndex, "000")
        If Dir(runFolder, vbDirectory) = "" Then Exit Do
        runIndex = runIndex + 1
    Loop

    MkDir runFolder
    MkDir runFolder & "\images"

    BASE_RUN_FOLDER = runFolder
    IMAGE_FOLDER = runFolder & "\images"
    OUTPUT_FOLDER = runFolder & "\output.json"
    LOG_FILE_PATH = runFolder & "\pdf_processing.log"

End Sub

Public Sub WriteLog(message As String, Optional level As LogLevel = LogLevel.Info)
    On Error Resume Next
    
    Dim fileNumber As Integer
    Dim timestamp As String
    Dim levelText As String
    Dim logEntry As String

    timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")

    Select Case level
        Case LogLevel.Info
            levelText = "INFO"
        Case LogLevel.Warning
            levelText = "WARNING"
        Case LogLevel.Error
            levelText = "ERROR"
        Case LogLevel.Debugging
            levelText = "DEBUG"
        Case Else
            levelText = "INFO"
    End Select

    logEntry = timestamp & " [" & levelText & "] " & message

    fileNumber = FreeFile

    Open LOG_FILE_PATH For Append As #fileNumber

    Print #fileNumber, logEntry

    Close #fileNumber

End Sub

Public Sub LogInfo(message As String)
    WriteLog message, LogLevel.Info
End Sub

Public Sub LogWarning(message As String)
    WriteLog message, LogLevel.Warning
End Sub

Public Sub LogError(message As String)
    WriteLog message, LogLevel.Error
End Sub

Public Sub LogDebug(message As String)
    WriteLog message, LogLevel.Debugging
End Sub

Public Sub LogSessionStart()
    WriteLog String(80, "="), LogLevel.Info
    WriteLog "PDF Processing Session Started", LogLevel.Info
    WriteLog String(80, "="), LogLevel.Info
End Sub

Public Sub LogSessionEnd()
    WriteLog String(80, "="), LogLevel.Info
    WriteLog "PDF Processing Session Ended", LogLevel.Info
    WriteLog String(80, "="), LogLevel.Info
    WriteLog "", LogLevel.Info
End Sub

Public Sub ClearLogFile()
    On Error Resume Next
    
    Dim fileNumber As Integer
    fileNumber = FreeFile

    Open LOG_FILE_PATH For Output As #fileNumber
    Close #fileNumber
    
    WriteLog "Log file cleared", LogLevel.Info
End Sub

Public Sub OpenLogFile()
    On Error Resume Next
    Shell "notepad.exe " & LOG_FILE_PATH, vbNormalFocus
End Sub

Public Sub PDFToExcel(control As IRibbonControl)

    InitRunFolders

    Dim pdfPath As String
    Dim imageFolder As String
    Dim shellCmd As String
    Dim wsh As Object
    Dim exitCode As Long

    LogInfo "Starting PDF to Excel conversion process"
    
    pdfPath = GetPDFFile()
    If pdfPath = "" Then
        LogWarning "No PDF file selected."
        Exit Sub
    End If

    imageFolder = IMAGE_FOLDER

    LogDebug "PDF path: " & pdfPath
    LogDebug "Image folder: " & imageFolder

    If Dir(imageFolder, vbDirectory) = "" Then
        MkDir imageFolder
        LogInfo "Created image folder: " & imageFolder
    Else
        LogDebug "Image folder already exists."
    End If

    shellCmd = "cmd /c cd /d """ & imageFolder & """ && pdftoppm -jpeg """ & pdfPath & """ page"
    LogDebug "Shell command: " & shellCmd

    Set wsh = CreateObject("WScript.Shell")
    LogInfo "Converting PDF to images using pdftoppm..."

    exitCode = wsh.Run(shellCmd, 0, True)
    LogInfo "pdftoppm process completed with exit code: " & exitCode

    If exitCode = 0 Then
        LogInfo "PDF to images conversion successful"
    Else
        LogError "Error converting PDF to images. Exit code: " & exitCode
        MsgBox "Error converting PDF to images. Exit code: " & exitCode, vbCritical
    End If

    Set wsh = Nothing
    
    LogInfo "Starting table extraction from images..."
    ConvertPDFToExcel (imageFolder)
    
    LogSessionEnd
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

    LogInfo "Starting ConvertPDFToExcel process"
    
    apiKey = "AIzaSyDRzpEHgPS7LRqDRmPx-mEXY-Cukyqr-o4"

    imageFolder = IMAGE_FOLDER
    
    If Right(imageFolder, 1) <> "\" Then
        imageFolder = imageFolder & "\"
    End If
    
    LogDebug "Final image folder path: " & imageFolder

    Application.StatusBar = "Extracting tables from images..."
    Application.ScreenUpdating = False
    
    LogInfo "Calling Gemini API to extract tables from images..."
    tableData = ExtractTablesWithGeminiFromImages(imageFolder, apiKey)
    Dim filePath As String
    Dim fileNumber As Integer
    
    filePath = OUTPUT_FOLDER
    LogDebug "Saving extracted data to: " & filePath
    
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, tableData
    Close #fileNumber
    
    LogInfo "Table data saved to output file"
    LogInfo "Starting Excel sheet population..."
    
    TestParseFromFile (filePath)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    LogInfo "ConvertPDFToExcel process completed"
End Sub

Private Function ExtractTablesWithGeminiFromImages(imageFolder As String, apiKey As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim combinedData As String
    Dim pageNum As Integer
    Dim imageCount As Integer
    Dim tableCount As Integer

    LogInfo "Starting table extraction from images in folder: " & imageFolder

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Dir(imageFolder, vbDirectory) = "" Then
        LogError "Selected folder does not exist or is inaccessible: " & imageFolder
        MsgBox "Selected folder does not exist or is inaccessible.", vbCritical
        Exit Function
    End If

    Set folder = fso.GetFolder(imageFolder)
    If folder Is Nothing Then
        LogError "Failed to access folder."
        MsgBox "Failed to access folder.", vbCritical
        Exit Function
    End If

    imageCount = 0
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Then
            imageCount = imageCount + 1
        End If
    Next file

    LogInfo "Found " & imageCount & " image files to process"

    pageNum = 1
    combinedData = "["

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Then
            LogInfo "Processing page " & pageNum & " of " & imageCount & ": " & file.Name
            Dim pageResponse As String
            pageResponse = ProcessImagePageWithGemini(file.Path, apiKey)

            If pageResponse <> "" Then
                ' Clean line breaks
                pageResponse = Trim(pageResponse)
                If Left(pageResponse, 1) = "[" And Right(pageResponse, 1) = "]" Then
                    pageResponse = Mid(pageResponse, 2, Len(pageResponse) - 2)
                End If

                If tableCount > 0 Then
                    combinedData = combinedData & "," & vbCrLf
                End If

                combinedData = combinedData & pageResponse
                tableCount = tableCount + 1

                LogInfo "Successfully extracted data from page " & pageNum
            Else
                LogWarning "No data extracted from page " & pageNum & ": " & file.Name
            End If
            pageNum = pageNum + 1
        End If
    Next file

    combinedData = combinedData & "]"

    LogInfo "Completed processing all images. Total pages processed: " & (pageNum - 1)
    ExtractTablesWithGeminiFromImages = combinedData
End Function

Private Function ProcessImagePageWithGemini(imagePath As String, apiKey As String) As String
    Dim fileUri As String
    Dim jsonRequest As String
    Dim responseText As String
    Dim http As Object
    
    LogDebug "Starting Gemini API processing for image: " & imagePath
    
    LogDebug "Uploading file to Gemini..."
    fileUri = UploadFileToGemini(imagePath, apiKey)
    If fileUri = "" Then
        LogError "Failed to upload file to Gemini: " & imagePath
        Exit Function
    End If
    
    LogDebug "File uploaded successfully. URI: " & fileUri
    
    jsonRequest = CreateGeminiImageRequest(fileUri)
    LogDebug "Created Gemini API request"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 60000, 60000, 60000, 300000
    http.Open "POST", "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" & apiKey, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    LogDebug "Sending request to Gemini API..."
    http.Send jsonRequest

    responseText = http.responseText
    LogDebug "Received response from Gemini API. Status: " & http.Status
    
    If http.Status = 200 Then
        LogInfo "Gemini API call successful for image: " & imagePath
        ProcessImagePageWithGemini = ParseGeminiResponse(responseText)
    Else
        LogError "Gemini API Error for image " & imagePath & ": " & http.Status & " - " & http.StatusText
        LogError "Response: " & responseText
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

    promptText = "Extract all tables from this image. Respond with ONLY a raw JSON array in this EXACT format: " & _
        "[{" & Chr(34) & "headers" & Chr(34) & ": [" & Chr(34) & "header1" & Chr(34) & ", " & Chr(34) & "header2" & Chr(34) & ", ...], " & _
        Chr(34) & "rows" & Chr(34) & ": [[" & Chr(34) & "value1" & Chr(34) & ", " & Chr(34) & "value2" & Chr(34) & ", ...], ...]}]. " & _
        "Each table in the image must be represented as a separate object in the array." & _
        "It is MANDATORY that every object contains BOTH a 'headers' array and a 'rows' array. Do NOT return tables without headers or rows." & _
        "Ensure all column headers are extracted and aligned properly. Do NOT invent data. If no headers are visible, infer logical placeholders like Column1, Column2." & _
        "Avoid markdown formatting, explanations, or code blocks. Only return raw JSON text." & _
        "If any column contains a signature field (e.g., 'User Sign', 'Signature', etc.), replace the value in that cell with 'SIGNATURE DETECTED' for each affected row." & _
        "Before finalizing the response, VERIFY that all rows match the headers in column count and meaning. Fix any misalignments or structural errors." & _
        "Again, respond with ONLY a valid JSON array with each object containing both 'headers' and 'rows'. No extra output." & _
        "VERY IMPORTANT: Do not wrap the output in code blocks, triple backticks, or markdown formatting. Return only raw JSON text without any surrounding characters."

    ' Escape for JSON compatibility
    promptText = Replace(promptText, "\", "\\")
    promptText = Replace(promptText, """", "\""")
    promptText = Replace(promptText, vbCrLf, "\n")
    promptText = Replace(promptText, vbLf, "\n")
    promptText = Replace(promptText, vbCr, "\n")
    
    ' Build JSON request
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

Private Function ParseGeminiResponse(response As String) As String
    On Error GoTo ErrorHandler

    Dim textStart As Long, textEnd As Long
    Dim searchPattern As String
    searchPattern = """text"": """

    textStart = InStr(response, searchPattern)
    If textStart = 0 Then
        ParseGeminiResponse = ""
        Exit Function
    End If

    textStart = textStart + Len(searchPattern)

    Dim pos As Long, inEscape As Boolean
    pos = textStart
    inEscape = False

    ' Find the closing quote of the "text" field (skipping escaped quotes)
    Do While pos <= Len(response)
        Dim currentChar As String
        currentChar = Mid(response, pos, 1)

        If currentChar = "\" And Not inEscape Then
            inEscape = True
        ElseIf currentChar = """" And Not inEscape Then
            textEnd = pos - 1
            Exit Do
        Else
            inEscape = False
        End If
        pos = pos + 1
    Loop

    If textEnd <= textStart Then
        ParseGeminiResponse = ""
        Exit Function
    End If

    ' Extract and clean the text
    Dim extractedText As String
    extractedText = Mid(response, textStart, textEnd - textStart + 1)

    extractedText = Replace(extractedText, "\n", "")
    extractedText = Replace(extractedText, "\r", "")
    extractedText = Replace(extractedText, "\t", "")
    extractedText = Replace(extractedText, "\""", """")
    extractedText = Replace(extractedText, "\\", "")
    extractedText = Replace(extractedText, "```json", "")
    extractedText = Replace(extractedText, "```", "")
    extractedText = Replace(extractedText, "\", "")
    extractedText = Replace(extractedText, "\", "/")
    extractedText = Replace(extractedText, """""", """")
    extractedText = Trim(extractedText)


    ' Now extract individual {...} blocks and rebuild the array
    Dim jsonOutput As String
    Dim objectStart As Long, objectEnd As Long
    Dim depth As Long
    Dim i As Long
    Dim insideObject As Boolean

    For i = 1 To Len(extractedText)
        Dim ch As String
        ch = Mid(extractedText, i, 1)

        If ch = "{" Then
            If depth = 0 Then objectStart = i
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                objectEnd = i
                Dim obj As String
                obj = Mid(extractedText, objectStart, objectEnd - objectStart + 1)
                If Len(jsonOutput) > 1 Then jsonOutput = jsonOutput & ","
                jsonOutput = jsonOutput & obj
            End If
        End If
    Next i

    ParseGeminiResponse = jsonOutput
    Exit Sub

ErrorHandler:
    ParseGeminiResponse = ""
End Sub

Private Sub TestParseFromFile(filePath As String)
    Dim fileContent As String
    Dim fileNumber As Integer
    
    LogInfo "Starting TestParseFromFile for: " & filePath

    If Dir(filePath) = "" Then
        LogError "File not found: " & filePath
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContent = Input(LOF(fileNumber), fileNumber)
    Close #fileNumber
    
    LogDebug "File content length: " & Len(fileContent) & " characters"

    If fileContent <> "" Then
        LogInfo "Parsing extracted data into Excel sheets..."
        Dim tableCount As Integer
        tableCount = ParseGeminiDataToSeparateSheets(fileContent)
        LogInfo "Parsing completed successfully! Created " & tableCount & " sheets for " & tableCount & " tables"
        MsgBox "Test parsing completed successfully! Created " & tableCount & " sheets for " & tableCount & " tables.", vbInformation
    Else
        LogError "File is empty or could not be read: " & filePath
        MsgBox "File is empty or could not be read.", vbCritical
    End If
End Sub

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

Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim protectedNames As Variant
    protectedNames = Array("Dashboard", "Summary", "Charts") ' Add your custom sheets here

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

Private Function ParseGeminiDataToSeparateSheets(jsonData As String) As Integer
    On Error GoTo ErrorHandler
    
    Dim tableCount As Integer
    tableCount = 0
    
    Dim tableStart As Long, tableEnd As Long
    Dim pos As Long
    
    LogInfo "Starting ParseGeminiDataToSeparateSheets"
    
    pos = 1
    
    Do While pos < Len(jsonData)
        tableStart = InStr(pos, jsonData, "{")
        If tableStart = 0 Then Exit Do

        tableEnd = FindObjectEnd(jsonData, tableStart)
        If tableEnd = 0 Then Exit Do
        
        Dim tableContent As String
        tableContent = Mid(jsonData, tableStart, tableEnd - tableStart + 1)

        tableCount = tableCount + 1
        LogInfo "Page " & tableCount & " converted to Excel sheet"
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
    LogError "Error parsing Gemini data to separate sheets: " & Err.Description
    MsgBox "Error parsing Gemini data to separate sheets: " & Err.Description, vbCritical
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