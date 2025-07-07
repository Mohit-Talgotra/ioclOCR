Public ribbon As IRibbonUI

Public Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub

Public Sub GenerateDSOCharts_RegionAndSubregions(control As IRibbonControl)
    Call ConsolidateByRegion

    Dim dataWS As Worksheet, chartWS As Worksheet
    Dim lastRow As Long, i As Long
    Dim subregionDict As Object, dsoCoverageDict As Object, dsoItemTypeDict As Object
    Dim itemTypeDict As Object, coverageDict As Object
    Dim subregion As String, item As String, amcVal As String, warrantyVal As String, coverageStatus As String
    Dim chartRow As Long
    Dim chartObj1 As ChartObject, chartObj2 As ChartObject
    Dim subregionKey As Variant, key As Variant

    Set dataWS = ActiveWorkbook.Sheets("Consolidated")
    lastRow = dataWS.Cells(dataWS.rows.Count, 1).End(xlUp).Row

    Set dsoCoverageDict = CreateObject("Scripting.Dictionary")
    Set dsoItemTypeDict = CreateObject("Scripting.Dictionary")
    Set subregionDict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        subregion = Trim(dataWS.Cells(i, 2).Value)
        item = UCase(Trim(dataWS.Cells(i, 1).Value))
        amcVal = UCase(Trim(dataWS.Cells(i, 4).Value))
        warrantyVal = UCase(Trim(dataWS.Cells(i, 5).Value))

        If amcVal = "YES" Then
            coverageStatus = "AMC"
        ElseIf warrantyVal = "YES" Then
            coverageStatus = "Warranty"
        Else
            coverageStatus = "Not Covered"
        End If

        dsoItemTypeDict(item) = dsoItemTypeDict(item) + 1
        dsoCoverageDict(coverageStatus) = dsoCoverageDict(coverageStatus) + 1

        If Not subregionDict.Exists(subregion) Then
            Set itemTypeDict = CreateObject("Scripting.Dictionary")
            Set coverageDict = CreateObject("Scripting.Dictionary")
            subregionDict(subregion) = Array(itemTypeDict, coverageDict)
        Else
            Set itemTypeDict = subregionDict(subregion)(0)
            Set coverageDict = subregionDict(subregion)(1)
        End If

        itemTypeDict(item) = itemTypeDict(item) + 1
        coverageDict(coverageStatus) = coverageDict(coverageStatus) + 1
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("DSO_Overview").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set chartWS = ActiveWorkbook.Sheets.Add
    chartWS.Name = "DSO_Overview"

    chartWS.Cells(1, 1).Value = "Coverage"
    chartWS.Cells(1, 3).Value = "Item Type"
    chartWS.Cells(1, 4).Value = "Count"
    chartWS.Range("A1:D1").Font.Color = RGB(255, 255, 255)
    chartRow = 2
    
    For Each key In dsoCoverageDict.Keys
        chartWS.Cells(chartRow, 1).Value = key
        chartWS.Cells(chartRow, 2).Value = dsoCoverageDict(key)
        chartWS.Range("A" & chartRow & ":B" & chartRow).Font.Color = RGB(255, 255, 255)
        chartRow = chartRow + 1
    Next key
    Dim dsoCoverageEnd As Long: dsoCoverageEnd = chartRow - 1

    chartRow = 2
    For Each key In dsoItemTypeDict.Keys
        chartWS.Cells(chartRow, 3).Value = key
        chartWS.Cells(chartRow, 4).Value = dsoItemTypeDict(key)
        chartWS.Range("C" & chartRow & ":D" & chartRow).Font.Color = RGB(255, 255, 255)
        chartRow = chartRow + 1
    Next key
    Dim dsoItemEnd As Long: dsoItemEnd = chartRow - 1

    Set chartObj1 = chartWS.ChartObjects.Add(Left:=10, Width:=300, Top:=10, Height:=250)
    With chartObj1.Chart
        .ChartType = xlPie
        .SetSourceData Source:=chartWS.Range("A2:B" & dsoCoverageEnd)
        .SeriesCollection(1).XValues = chartWS.Range("A2:A" & dsoCoverageEnd)
        .SeriesCollection(1).Values = chartWS.Range("B2:B" & dsoCoverageEnd)
        .HasTitle = True
        .ChartTitle.text = "Coverage - DSO"
        .ApplyDataLabels
    End With
    
    Set chartObj2 = chartWS.ChartObjects.Add(Left:=330, Width:=300, Top:=10, Height:=250)
    With chartObj2.Chart
        .ChartType = xlPie
        .SetSourceData Source:=chartWS.Range("C2:D" & dsoItemEnd)
        .SeriesCollection(1).XValues = chartWS.Range("C2:C" & dsoItemEnd)
        .SeriesCollection(1).Values = chartWS.Range("D2:D" & dsoItemEnd)
        .HasTitle = True
        .ChartTitle.text = "Asset Types - DSO"
        .ApplyDataLabels
    End With

    Dim btn As Shape
    Set btn = chartWS.Shapes.AddShape(msoShapeRectangle, 240, 270, 180, 30)
    With btn
        .Name = "btnShowRegions"
        .TextFrame2.TextRange.text = "Show Region Charts"
        .OnAction = "ShowRegionCharts"
        .Fill.ForeColor.RGB = RGB(40, 150, 255)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With

    Dim chartTop As Long: chartTop = 320
    For Each subregionKey In subregionDict.Keys
        Set itemTypeDict = subregionDict(subregionKey)(0)
        Set coverageDict = subregionDict(subregionKey)(1)
    
        chartWS.Cells(chartTop \ 20, 1).Value = "Coverage"
        chartWS.Cells(chartTop \ 20, 3).Value = "Item Type"
        chartWS.Cells(chartTop \ 20, 4).Value = "Count"
        chartWS.Range("A" & (chartTop \ 20) & ":D" & (chartTop \ 20)).Font.Color = RGB(255, 255, 255)
    
        chartRow = (chartTop \ 20) + 1
        For Each key In coverageDict.Keys
            chartWS.Cells(chartRow, 1).Value = key
            chartWS.Cells(chartRow, 2).Value = coverageDict(key)
            chartWS.Range("A" & chartRow & ":B" & chartRow).Font.Color = RGB(255, 255, 255)
            chartRow = chartRow + 1
        Next key
        Dim covStart As Long: covStart = (chartTop \ 20) + 1
        Dim covEnd As Long: covEnd = chartRow - 1
    
        chartRow = (chartTop \ 20) + 1
        For Each key In itemTypeDict.Keys
            chartWS.Cells(chartRow, 3).Value = key
            chartWS.Cells(chartRow, 4).Value = itemTypeDict(key)
            chartWS.Range("C" & chartRow & ":D" & chartRow).Font.Color = RGB(255, 255, 255)
            chartRow = chartRow + 1
        Next key
        Dim itmStart As Long: itmStart = (chartTop \ 20) + 1
        Dim itmEnd As Long: itmEnd = chartRow - 1
    
        Set chartObj1 = chartWS.ChartObjects.Add(Left:=10, Width:=300, Top:=chartTop, Height:=250)
        With chartObj1
            .Name = "Subregion_" & subregionKey & "_Coverage"
            .Visible = False
            With .Chart
                .ChartType = xlPie
                .SetSourceData chartWS.Range("A" & covStart & ":B" & covEnd)
                .SeriesCollection(1).XValues = chartWS.Range("A" & covStart & ":A" & covEnd)
                .SeriesCollection(1).Values = chartWS.Range("B" & covStart & ":B" & covEnd)
                .HasTitle = True
                .ChartTitle.text = "Coverage - " & subregionKey
                .ApplyDataLabels
            End With
        End With
    
        Set chartObj2 = chartWS.ChartObjects.Add(Left:=330, Width:=300, Top:=chartTop, Height:=250)
        With chartObj2
            .Name = "Subregion_" & subregionKey & "_Items"
            .Visible = False
            With .Chart
                .ChartType = xlPie
                .SetSourceData chartWS.Range("C" & itmStart & ":D" & itmEnd)
                .SeriesCollection(1).XValues = chartWS.Range("C" & itmStart & ":C" & itmEnd)
                .SeriesCollection(1).Values = chartWS.Range("D" & itmStart & ":D" & itmEnd)
                .HasTitle = True
                .ChartTitle.text = "Asset Types - " & subregionKey
                .ApplyDataLabels
            End With
        End With
    
        chartTop = chartTop + 280
    Next subregionKey

    MsgBox "Charts generated successfully.", vbInformation
End Sub

Private Sub ConsolidateByRegion()

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim colAsset As Long, colSign As Long, colAMC As Long, colRegion As Long, colWarranty As Long
    Dim c As Range
    Dim destWs As Worksheet

    On Error Resume Next
    Set destWs = ActiveWorkbook.Worksheets("Consolidated")
    If destWs Is Nothing Then
        Set destWs = ActiveWorkbook.Worksheets.Add
        destWs.Name = "Consolidated"
    Else
        destWs.Cells.ClearContents
    End If
    On Error GoTo 0

    destWs.Range("A1:F1").Value = Array("Asset Type", "Region", "Status", "AMC", "Warranty", "Source Link")
    Dim outRow As Long: outRow = 2

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Dashboard" And ws.Name <> "Consolidated" And ws.Name <> "DSO_Overview" Then

            colAsset = 0: colSign = 0: colAMC = 0: colRegion = 0: colWarranty = 0

            For Each c In ws.rows(1).Cells
                If colAsset = 0 Then
                    Dim colText As String
                    colText = Trim(LCase(c.text))
                    If colText = "asset" Or colText = "asset type" Then
                        colAsset = c.Column
                    End If
                End If
                If colSign = 0 And LCase(c.text) Like "*sign*" Then colSign = c.Column
                If colAMC = 0 And (LCase(c.text) Like "*amc*" Or LCase(c.text) Like "*contract*") Then colAMC = c.Column
                If colWarranty = 0 And LCase(c.text) Like "*warranty*" Then colWarranty = c.Column
                If colRegion = 0 And LCase(c.text) Like "*region*" Then colRegion = c.Column
            Next c

            If colAsset > 0 And colSign > 0 And colRegion > 0 Then
                lastRow = ws.Cells(ws.rows.Count, colAsset).End(xlUp).Row
                For r = 2 To lastRow
                    If Application.WorksheetFunction.CountA(ws.rows(r)) > 0 Then
                        Dim assetType As String, signVal As String, amcVal As String, warrantyVal As String, regionVal As String
                        Dim rawAsset As String
                        rawAsset = Trim(LCase(ws.Cells(r, colAsset).text))
                        
                        Select Case rawAsset
                            Case "switch", "switches"
                                assetType = "SWITCH"
                            Case "router", "routers"
                                assetType = "ROUTER"
                            Case "printer", "printers"
                                assetType = "PRINTER"
                            Case "monitor", "monitors"
                                assetType = "MONITOR"
                            Case "desktop", "desktops", "pc", "pcs"
                                assetType = "DESKTOP"
                            Case "laptop", "laptops"
                                assetType = "LAPTOP"
                            Case "polycom", "polycom camera"
                                assetType = "WEBCAM"
                            Case Else
                                assetType = UCase(rawAsset)
                        End Select

                        signVal = Trim(ws.Cells(r, colSign).text)
                        regionVal = Trim(ws.Cells(r, colRegion).text)
                        amcVal = "No": warrantyVal = "No"

                        If colAMC > 0 Then
                            If Not IsError(ws.Cells(r, colAMC)) Then
                                If LCase(Trim(ws.Cells(r, colAMC).text)) = "amc" Then amcVal = "Yes"
                            End If
                        End If

                        If colWarranty > 0 Then
                            If Not IsError(ws.Cells(r, colWarranty)) Then
                                If LCase(Trim(ws.Cells(r, colWarranty).text)) = "warranty" Then warrantyVal = "Yes"
                            End If
                        End If

                        If Len(assetType) > 0 Then
                            destWs.Cells(outRow, 1).Value = assetType
                            destWs.Cells(outRow, 2).Value = regionVal
                            destWs.Cells(outRow, 3).Value = IIf(UCase(signVal) = "SIGNATURE DETECTED", "Working", "Defective")
                            destWs.Cells(outRow, 4).Value = amcVal
                            destWs.Cells(outRow, 5).Value = warrantyVal
                            
                            Dim linkAddress As String, linkText As String
                            linkAddress = "'" & ws.Name & "'!" & ws.Cells(r, colAsset).Address
                            linkText = "Go to Cell"
                            destWs.Hyperlinks.Add Anchor:=destWs.Cells(outRow, 6), _
                                Address:="", _
                                SubAddress:=linkAddress, _
                                TextToDisplay:=linkText
                            
                            outRow = outRow + 1
                        End If
                    End If
                Next r
            End If
        End If
    Next ws

    MsgBox "Consolidated sheet generated.", vbInformation

End Sub

Public Sub ShowRegionCharts()
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("DSO_Overview")

    For Each chartObj In ws.ChartObjects
        If InStr(1, chartObj.Name, "Subregion_") > 0 Then
            chartObj.Visible = True
        End If
    Next chartObj

    MsgBox "Region charts revealed.", vbInformation
End Sub