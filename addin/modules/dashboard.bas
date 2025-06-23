Option Explicit

Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub ConsolidateByRegion(control As IRibbonControl)

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim colAsset As Long, colSign As Long, colAMC As Long, colRegion As Long
    Dim c As Range
    Dim destWs As Worksheet

    On Error Resume Next
    Set destWs = Worksheets("Consolidated")
    If destWs Is Nothing Then
        Set destWs = Worksheets.Add
        destWs.Name = "Consolidated"
    Else
        destWs.Cells.ClearContents
    End If
    On Error GoTo 0

    destWs.Range("A1:D1").Value = Array("Asset Type", "Region", "Status", "AMC")
    Dim outRow As Long: outRow = 2

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" And ws.Name <> "Consolidated" Then

            colAsset = 0: colSign = 0: colAMC = 0: colRegion = 0

            For Each c In ws.rows(1).Cells
                If colAsset = 0 And (LCase(c.text) Like "*asset*" Or LCase(c.text) Like "*type*") Then colAsset = c.Column
                If colSign = 0 And LCase(c.text) Like "*sign*" Then colSign = c.Column
                If colAMC = 0 And (LCase(c.text) Like "*warranty*" Or LCase(c.text) Like "*contract*" Or LCase(c.text) Like "*amc*") Then colAMC = c.Column
                If colRegion = 0 And LCase(c.text) Like "*region*" Then colRegion = c.Column
            Next c

            If colAsset > 0 And colSign > 0 And colRegion > 0 Then
                lastRow = ws.Cells(ws.rows.Count, colAsset).End(xlUp).Row
                For r = 2 To lastRow
                    If Application.WorksheetFunction.CountA(ws.rows(r)) > 0 Then
                        Dim assetType As String, signVal As String, amcVal As String, regionVal As String
                        assetType = Trim(ws.Cells(r, colAsset).text)
                        signVal = Trim(ws.Cells(r, colSign).text)
                        regionVal = Trim(ws.Cells(r, colRegion).text)

                        If colAMC > 0 Then
                            amcVal = Trim(ws.Cells(r, colAMC).text)
                        Else
                            amcVal = "NO"
                        End If

                        If Len(assetType) > 0 Then
                            destWs.Cells(outRow, 1).Value = assetType
                            destWs.Cells(outRow, 2).Value = regionVal
                            destWs.Cells(outRow, 3).Value = IIf(UCase(signVal) = "SIGNATURE DETECTED", "Working", "Defective")
                            destWs.Cells(outRow, 4).Value = IIf(UCase(amcVal) = "AMC", "Yes", "No")
                            outRow = outRow + 1
                        End If
                    End If
                Next r
            Else
                Debug.Print "Skipped sheet '" & ws.Name & "' — missing required columns:"
                If colAsset = 0 Then Debug.Print "  > Asset Type or Type"
                If colSign = 0 Then Debug.Print "  > User Sign"
                If colRegion = 0 Then Debug.Print "  > Region"
            End If

        End If
    Next ws

    MsgBox "Consolidated sheet generated.", vbInformation

End Sub
