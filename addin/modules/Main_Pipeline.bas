Public BASE_RUN_FOLDER As String
Private IMAGE_FOLDER As String
Private OUTPUT_FOLDER As String
Public LOG_FILE_PATH As String

Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub PDFToExcel(control As IRibbonControl)
    InitRunFolders

    Call Logging.LogSessionStart

    Dim pdfPath As String
    Dim imageFolder As String
    Dim shellCmd As String
    Dim wsh As Object
    Dim exitCode As Long

    Call Logging.LogInfo("Starting PDF to Excel conversion process")
    
    pdfPath = GetPDFFile()
    If pdfPath = "" Then
        Call Logging.LogWarning("No PDF file selected.")
        Exit Sub
    End If

    imageFolder = IMAGE_FOLDER

    Call Logging.LogDebug("PDF path: " & pdfPath)
    Call Logging.LogDebug("Image folder: " & imageFolder)

    If Dir(imageFolder, vbDirectory) = "" Then
        MkDir imageFolder
        Call Logging.LogInfo("Created image folder: " & imageFolder)
    Else
        Call Logging.LogDebug("Image folder already exists.")
    End If

    shellCmd = "cmd /c cd /d """ & imageFolder & """ && pdftoppm -jpeg """ & pdfPath & """ page"
    Call Logging.LogDebug("Shell command: " & shellCmd)

    Set wsh = CreateObject("WScript.Shell")
    Call Logging.LogInfo("Converting PDF to images using pdftoppm...")

    exitCode = wsh.Run(shellCmd, 0, True)
    Call Logging.LogInfo("pdftoppm process completed with exit code: " & exitCode)

    If exitCode = 0 Then
        Call Logging.LogInfo("PDF to images conversion successful")
    Else
        Call Logging.LogError("Error converting PDF to images. Exit code: " & exitCode)
        MsgBox "Error converting PDF to images. Exit code: " & exitCode, vbCritical
    End If

    Set wsh = Nothing
    
    Call Logging.LogInfo("Starting table extraction from images...")
    ConvertPDFToExcel (imageFolder)
    
    Call Logging.LogSessionEnd
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
    
    apiKey = "Insert GEMINI API Key Here"

    imageFolder = IMAGE_FOLDER
    
    If Right(imageFolder, 1) <> "\" Then
        imageFolder = imageFolder & "\"
    End If
    
    Call Logging.LogDebug("Final image folder path: " & imageFolder)

    Application.StatusBar = "Extracting tables from images..."
    Application.ScreenUpdating = False
    
    Call Logging.LogInfo("Calling Gemini API to extract tables from images...")
    tableData = Gemini.ExtractTablesWithGeminiFromImages(imageFolder, apiKey)
    Dim filePath As String
    Dim fileNumber As Integer
    
    filePath = OUTPUT_FOLDER
    Call Logging.LogDebug("Saving extracted data to: " & filePath)
    
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, tableData
    Close #fileNumber
    
    Call Logging.LogInfo("Table data saved to output file")
    Call Logging.LogInfo("Starting Excel sheet population...")
    
    TestParseFromFile (filePath)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Call Logging.LogInfo("ConvertPDFToExcel process completed")
End Sub

Private Sub TestParseFromFile(filePath As String)
    Dim fileContent As String
    Dim fileNumber As Integer
    
    Call Logging.LogInfo("Starting TestParseFromFile for: " & filePath)

    If Dir(filePath) = "" Then
        Call Logging.LogError("File not found: " & filePath)
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContent = Input(LOF(fileNumber), fileNumber)
    Close #fileNumber
    
    Call Logging.LogDebug("File content length: " & Len(fileContent) & " characters")

    If fileContent <> "" Then
        Call Logging.LogInfo("Parsing extracted data into Excel sheets...")
        Dim tableCount As Integer
        tableCount = Parsing.ParseGeminiDataToSeparateSheets(fileContent)
        Call Logging.LogInfo("Parsing completed successfully! Created " & tableCount & " sheets for " & tableCount & " tables")
        MsgBox "Test parsing completed successfully! Created " & tableCount & " sheets for " & tableCount & " tables.", vbInformation
    Else
        Call Logging.LogError("File is empty or could not be read: " & filePath)
        MsgBox "File is empty or could not be read.", vbCritical
    End If
End Sub