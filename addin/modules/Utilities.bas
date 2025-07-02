Public Function FindObjectEnd(text As String, startPos As Long) As Long
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