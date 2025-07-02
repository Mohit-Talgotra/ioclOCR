Public Function ExtractTablesWithGeminiFromImages(imageFolder As String, apiKey As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim combinedData As String
    Dim pageNum As Integer
    Dim imageCount As Integer
    Dim tableCount As Integer

    Call Logging.LogInfo("Starting table extraction from images in folder: " & imageFolder)

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Dir(imageFolder, vbDirectory) = "" Then
        Call Logging.LogError("Selected folder does not exist or is inaccessible: " & imageFolder)
        MsgBox "Selected folder does not exist or is inaccessible.", vbCritical
        Exit Function
    End If

    Set folder = fso.GetFolder(imageFolder)
    If folder Is Nothing Then
        Call Logging.LogError("Failed to access folder.")
        MsgBox "Failed to access folder.", vbCritical
        Exit Function
    End If

    imageCount = 0
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Then
            imageCount = imageCount + 1
        End If
    Next file

    Call Logging.LogInfo("Found " & imageCount & " image files to process")

    pageNum = 1
    combinedData = "["

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Then
            Call Logging.LogInfo("Processing page " & pageNum & " of " & imageCount & ": " & file.Name)
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

                Call Logging.LogInfo("Successfully extracted data from page " & pageNum)
            Else
                Call Logging.LogWarning("No data extracted from page " & pageNum & ": " & file.Name)
            End If
            pageNum = pageNum + 1
        End If
    Next file

    combinedData = combinedData & "]"

    Call Logging.LogInfo("Completed processing all images. Total pages processed: " & (pageNum - 1))
    ExtractTablesWithGeminiFromImages = combinedData
End Function

Private Function ProcessImagePageWithGemini(imagePath As String, apiKey As String) As String
    Dim fileUri As String
    Dim jsonRequest As String
    Dim responseText As String
    Dim http As Object
    
    Call Logging.LogDebug("Starting Gemini API processing for image: " & imagePath)
    
    Call Logging.LogDebug("Uploading file to Gemini...")
    fileUri = UploadFileToGemini(imagePath, apiKey)
    If fileUri = "" Then
        Call Logging.LogError("Failed to upload file to Gemini: " & imagePath)
        Exit Function
    End If
    
    Call Logging.LogDebug("File uploaded successfully. URI: " & fileUri)
    
    jsonRequest = CreateGeminiImageRequest(fileUri)
    Call Logging.LogDebug("Created Gemini API request")

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 60000, 60000, 60000, 300000
    http.Open "POST", "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" & apiKey, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    Call Logging.LogDebug("Sending request to Gemini API...")
    http.Send jsonRequest

    responseText = http.responseText
    Call Logging.LogDebug("Received response from Gemini API. Status: " & http.Status)
    
    If http.Status = 200 Then
        Call Logging.LogInfo("Gemini API call successful for image: " & imagePath)
        ProcessImagePageWithGemini = ParseGeminiResponse(responseText)
    Else
        Call Logging.LogError("Gemini API Error for image " & imagePath & ": " & http.Status & " - " & http.StatusText)
        Call Logging.LogError("Response: " & responseText)
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

    promptText = "You are to extract ALL tables from the provided image and respond with ONLY a valid JSON array in this STRICT format: " & _
    "[{" & Chr(34) & "headers" & Chr(34) & ": [" & Chr(34) & "header1" & Chr(34) & ", " & Chr(34) & "header2" & Chr(34) & ", ...], " & _
    Chr(34) & "rows" & Chr(34) & ": [[" & Chr(34) & "value1" & Chr(34) & ", " & Chr(34) & "value2" & Chr(34) & ", ...], ...]}]" & vbCrLf & _
    "CRITICAL RULES:" & vbCrLf & _
    "- You MUST return only raw JSON, without any surrounding markdown, text, explanations, code fences, or extra quotation marks." & vbCrLf & _
    "- Every object in the array must contain BOTH a 'headers' key (array of strings) AND a 'rows' key (array of arrays)." & vbCrLf & _
    "- Do NOT use extra double quotes around the JSON block or incorrectly escape characters." & vbCrLf & _
    "- Validate your output is strict JSON and parseable — NO syntax errors, mismatched brackets, or trailing commas." & vbCrLf & _
    "- If headers are missing in the image, generate logical placeholders like 'Column1', 'Column2', etc." & vbCrLf & _
    "- If any cell represents a signature (e.g., column header is 'Signature' or 'User Sign'), replace that cell's value with 'SIGNATURE DETECTED'." & vbCrLf & _
    "- Ensure all rows align to the same number of columns as the headers — fix any column mismatch before responding." & vbCrLf & _
    "FINAL WARNING: Respond with ONLY a raw JSON array conforming to the format above. Do NOT add text before or after it. No markdown. No explanations. Just JSON."

    ' Escape for JSON payload
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
    
    Call Logging.SavePromptToFile(jsonRequest)
    
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
    ' extractedText = Replace(extractedText, """""", """")
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
    Exit Function

ErrorHandler:
    ParseGeminiResponse = ""
End Function
