Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub CalculateDelta(control As IRibbonControl)
    Dim inventoryPath As String
    Dim wbInv As Workbook
    Dim wsInv As Worksheet
    Dim totalInventoryAssets As Double
    Dim totalDSOAssets As Double
    Dim lastRow As Long, i As Long
    Dim quantityCol As Long, mainLineCol As Long
    Dim quantityVal As Variant, mainLineVal As String
    Dim wsDSO As Worksheet
    Dim resultWS As Worksheet

    ' === STEP 1: OPEN EXTERNAL INVENTORY FILE ===
    inventoryPath = "C:\Users\talgo\Downloads\Overall final contract.xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbInv = Workbooks.Open(fileName:=inventoryPath, ReadOnly:=True)

    For Each wsInv In wbInv.Sheets
        lastRow = wsInv.Cells(wsInv.rows.Count, 1).End(xlUp).Row

        quantityCol = 0: mainLineCol = 0
        For i = 1 To wsInv.Cells(1, wsInv.Columns.Count).End(xlToLeft).Column
            If LCase(Trim(wsInv.Cells(1, i).Value)) = "quantity" Then quantityCol = i
            If LCase(Trim(wsInv.Cells(1, i).Value)) = "main line short text" Then mainLineCol = i
        Next i

        If quantityCol = 0 Or mainLineCol = 0 Then GoTo NextSheet

        For i = 2 To lastRow
            mainLineVal = wsInv.Cells(i, mainLineCol).text
            If InStr(1, mainLineVal, "AMC 2024-27(1100/", vbTextCompare) > 0 Then
                quantityVal = wsInv.Cells(i, quantityCol).Value
                If IsNumeric(quantityVal) Then
                    totalInventoryAssets = totalInventoryAssets + (quantityVal / 1095)
                End If
            End If
        Next i
NextSheet:
    Next wsInv

    wbInv.Close SaveChanges:=False

    ' === STEP 2: READ ONLY THE FIRST Item Type–Count TABLE FROM DSO_Overview ===
    Set wsDSO = ActiveWorkbook.Sheets("DSO_Overview")

    Dim startRow As Long: startRow = 2
    Dim itemCol As Long: itemCol = 3 ' Column C
    Dim countCol As Long: countCol = 4 ' Column D

    Do While Len(Trim(wsDSO.Cells(startRow, itemCol).Value)) > 0
        If IsNumeric(wsDSO.Cells(startRow, countCol).Value) Then
            totalDSOAssets = totalDSOAssets + wsDSO.Cells(startRow, countCol).Value
        End If
        startRow = startRow + 1
    Loop

    ' === STEP 3: CREATE COMPARISON SHEET ===
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Inventory Comparison").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set resultWS = ActiveWorkbook.Sheets.Add
    resultWS.Name = "Inventory Comparison"

    With resultWS
        .Range("A1").Value = "Metric"
        .Range("B1").Value = "Value"
        .Range("A2").Value = "Total Assets (from DSO_Overview)"
        .Range("B2").Value = totalDSOAssets
        .Range("A3").Value = "Total Assets (from Inventory workbook)"
        .Range("B3").Value = totalInventoryAssets
        .Range("A4").Value = "Difference (Inventory - DSO)"
        .Range("B4").Value = totalInventoryAssets - totalDSOAssets
    End With

    Application.ScreenUpdating = True
    MsgBox "Comparison complete. See 'Inventory Comparison' sheet.", vbInformation
End Sub