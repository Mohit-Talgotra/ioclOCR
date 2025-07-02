Public Enum LogLevel
    Info = 1
    Warning = 2
    Error = 3
    Debugging = 4
End Enum

Public Sub WriteLog(message As String, LOG_FILE_PATH As String, Optional level As LogLevel = LogLevel.Info)
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
    WriteLog message, LOG_FILE_PATH, LogLevel.Info
End Sub

Public Sub LogWarning(message As String)
    WriteLog message, LOG_FILE_PATH, LogLevel.Warning
End Sub

Public Sub LogError(message As String)
    WriteLog message, LOG_FILE_PATH, LogLevel.Error
End Sub

Public Sub LogDebug(message As String)
    WriteLog message, LOG_FILE_PATH, LogLevel.Debugging
End Sub

Public Sub LogSessionStart()
    WriteLog String(80, "="), LOG_FILE_PATH, LogLevel.Info
    WriteLog "PDF Processing Session Started", LOG_FILE_PATH, LogLevel.Info
    WriteLog String(80, "="), LOG_FILE_PATH, LogLevel.Info
End Sub

Public Sub LogSessionEnd()
    WriteLog String(80, "="), LOG_FILE_PATH, LogLevel.Info
    WriteLog "PDF Processing Session Ended", LOG_FILE_PATH, LogLevel.Info
    WriteLog String(80, "="), LOG_FILE_PATH, LogLevel.Info
    WriteLog "", LOG_FILE_PATH, LogLevel.Info
End Sub

Public Sub OpenLogFile()
    On Error Resume Next
    Shell "notepad.exe " & LOG_FILE_PATH, vbNormalFocus
End Sub

Public Sub SavePromptToFile(promptText As String)
    Dim fNum As Integer
    Dim promptPath As String

    promptPath = BASE_RUN_FOLDER & "\prompt_log.txt"
    fNum = FreeFile

    Open promptPath For Output As #fNum
    Print #fNum, promptText
    Close #fNum

    LogInfo "Prompt written to: " & promptPath
End Sub
