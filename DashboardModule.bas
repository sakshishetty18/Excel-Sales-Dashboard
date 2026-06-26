Attribute VB_Name = "DashboardModule"
Option Explicit

Sub BuildSalesDashboard()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsDash As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRng As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("Data")

    ' STEP 1: Convert Data sheet range to Excel Table named tblSales
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Set dataRng = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    On Error Resume Next
    wsData.ListObjects("tblSales").Delete
    On Error GoTo 0

    Set lo = wsData.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
    lo.Name = "tblSales"
    lo.TableStyle = "TableStyleMedium2"

    ' STEP 2: Delete and recreate Pivot sheet
    On Error Resume Next
    wb.Sheets("Pivot").Delete
    On Error GoTo 0

    Set wsPivot = wb.Sheets.Add(After:=wsData)
    wsPivot.Name = "Pivot"

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.ListObjects("tblSales").Range)

    Call CreatePivot_Monthly(pc, wsPivot)
    Call CreatePivot_Region(pc, wsPivot)
    Call CreatePivot_Manager(pc, wsPivot)
    Call CreatePivot_Category(pc, wsPivot)
    Call CreatePivot_Segment(pc, wsPivot)
    Call CreatePivot_TopProducts(pc, wsPivot)
    Call CreatePivot_KPI(pc, wsPivot)

    ' STEP 3: Delete and recreate Dashboard sheet
    On Error Resume Next
    wb.Sheets("Dashboard").Delete
    On Error GoTo 0

    Set wsDash = wb.Sheets.Add(Before:=wsData)
    wsDash.Name = "Dashboard"

    wsDash.Activate
    ActiveWindow.DisplayGridlines = False
    ' NOTE: Zoom is applied at the END — NOT here

    ' Background and font
    wsDash.Cells.Interior.Color = RGB(242, 244, 248)
    wsDash.Cells.Font.Name = "Calibri"

    ' Column widths: A-R = 8.2, S = 0.6, T-V = 8.2
    Dim i As Integer
    For i = 1 To 18
        wsDash.Columns(i).ColumnWidth = 8.2
    Next i
    wsDash.Columns(19).ColumnWidth = 0.6
    For i = 20 To 22
        wsDash.Columns(i).ColumnWidth = 8.2
    Next i
    wsDash.Rows.RowHeight = 13.5
    wsDash.Rows(11).RowHeight = 0

    ' STEP 3b: Header Banner
    Dim hdrLeft   As Single: hdrLeft   = wsDash.Range("A1").Left
    Dim hdrTop    As Single: hdrTop    = wsDash.Range("A1").Top
    Dim hdrWidth  As Single: hdrWidth  = wsDash.Range("A1:V1").Width
    Dim hdrHeight As Single: hdrHeight = wsDash.Range("A1:A3").Height

    Dim hdrShape As Shape
    Set hdrShape = wsDash.Shapes.AddShape(msoShapeRectangle, _
        hdrLeft, hdrTop, hdrWidth, hdrHeight)
    With hdrShape
        .Fill.ForeColor.RGB = RGB(11, 118, 160)
        .Fill.Transparency  = 0
        .Line.Visible       = msoFalse
        .Name               = "HeaderBanner"
    End With
    With hdrShape.TextFrame
        .MarginLeft   = 0
        .MarginRight  = 0
        .MarginTop    = 0
        .MarginBottom = 0
        .Characters.Text           = "Sales & Target Performance Dashboard"
        .Characters.Font.Bold      = True
        .Characters.Font.Color     = RGB(255, 255, 255)
        .Characters.Font.Size      = 36
        .Characters.Font.Name      = "Calibri"
        .HorizontalAlignment       = xlHAlignCenter
        .VerticalAlignment         = xlVAlignCenter
    End With

    ' STEP 4: KPI Cards
    Call CreateKPICards(wsDash, wsPivot)

    ' STEP 5: Charts
    Call CreateCharts(wsDash, wsPivot)

    ' STEP 6: Slicers
    Call CreateSlicers(wb, wsDash, wsPivot)

    wsDash.Activate
    wsDash.Range("A1").Select

    ' STEP 7: Turn screen on, then apply zoom with force-render cycle
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    ActiveWindow.Zoom = 100
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

    ActiveWindow.Zoom = 70
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

    wb.RefreshAll

    Call InjectDashboardEvent(wb)
    Call UpdateKPIValues(wsDash, wsPivot)
    Call FixDonutCenter

    MsgBox "Sales Dashboard built successfully!", vbInformation, "Dashboard Ready"

End Sub

Private Sub InjectDashboardEvent(wb As Workbook)
    On Error Resume Next

    Dim cm As Object
    Dim vbc As Object
    For Each vbc In wb.VBProject.VBComponents
        If vbc.Properties("Name") = "Dashboard" Then
            Set cm = vbc.CodeModule
            Exit For
        End If
    Next vbc

    If cm Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    Dim i As Long
    Dim startLine As Long, endLine As Long
    For i = cm.CountOfLines To 1 Step -1
        Dim lineText As String
        lineText = cm.Lines(i, 1)
        If InStr(lineText, "Worksheet_Calculate") > 0 Or _
           InStr(lineText, "Worksheet_PivotTableUpdate") > 0 Then
            startLine = i
            endLine = i
            Do While endLine < cm.CountOfLines
                endLine = endLine + 1
                If InStr(cm.Lines(endLine, 1), "End Sub") > 0 Then Exit Do
            Loop
            cm.DeleteLines startLine, endLine - startLine + 1
        End If
    Next i

    Dim calcEvent As String
    calcEvent = "Private Sub Worksheet_Calculate()" & vbCrLf & _
                "    On Error Resume Next" & vbCrLf & _
                "    Dim wsPivot As Worksheet" & vbCrLf & _
                "    Set wsPivot = ThisWorkbook.Sheets(""Pivot"")" & vbCrLf & _
                "    If Not wsPivot Is Nothing Then" & vbCrLf & _
                "        Call DashboardModule.UpdateKPIValues(Me, wsPivot)" & vbCrLf & _
                "    End If" & vbCrLf & _
                "    On Error GoTo 0" & vbCrLf & _
                "End Sub"

    Dim pvtEvent As String
    pvtEvent = "Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)" & vbCrLf & _
               "    On Error Resume Next" & vbCrLf & _
               "    Dim wsPivot As Worksheet" & vbCrLf & _
               "    Set wsPivot = ThisWorkbook.Sheets(""Pivot"")" & vbCrLf & _
               "    If Not wsPivot Is Nothing Then" & vbCrLf & _
               "        Call DashboardModule.UpdateKPIValues(Me, wsPivot)" & vbCrLf & _
               "    End If" & vbCrLf & _
               "    On Error GoTo 0" & vbCrLf & _
               "End Sub"

    Dim insertAt As Long
    insertAt = cm.CountOfLines + 1
    cm.InsertLines insertAt, calcEvent
    cm.InsertLines cm.CountOfLines + 1, pvtEvent

    On Error GoTo 0
End Sub

Private Sub CreatePivot_Monthly(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A2"), _
        TableName:="pvtMonthly")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Month").Orientation = xlRowField
        .PivotFields("Month").Position = 1

        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Total Sales"
            .NumberFormat = "#,##0"
        End With

        With .PivotFields("Target Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Total Target"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Region(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A20"), _
        TableName:="pvtRegion")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Region").Position = 1

        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sales by Region"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Manager(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A35"), _
        TableName:="pvtManager")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Sales Manager").Orientation = xlRowField
        .PivotFields("Sales Manager").Position = 1

        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sales by Manager"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Category(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F2"), _
        TableName:="pvtCategory")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Product Category").Orientation = xlRowField
        .PivotFields("Product Category").Position = 1

        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sales by Category"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_Segment(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F20"), _
        TableName:="pvtSegment")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        .PivotFields("Customer Segment").Orientation = xlRowField
        .PivotFields("Customer Segment").Position = 1

        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sales by Segment"
            .NumberFormat = "#,##0"
        End With

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_TopProducts(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    Dim pf As PivotField
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("F35"), _
        TableName:="pvtTopProducts")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        Set pf = .PivotFields("Product Name")
        pf.Orientation = xlRowField
        pf.Position = 1

        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Top Product Sales"
            .NumberFormat = "#,##0"
        End With

        On Error Resume Next
        pf.AutoShow 1, 1, 10, "Top Product Sales"
        On Error GoTo 0

        .RowAxisLayout xlTabularRow
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium2"
        .DisplayFieldCaptions = False
    End With
End Sub

Private Sub CreatePivot_KPI(pc As PivotCache, ws As Worksheet)
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("K5"), _
        TableName:="PT_KPI")
    On Error GoTo 0
    If pt Is Nothing Then Exit Sub

    With pt
        With .PivotFields("Sales Amount")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Total Sales"
            .NumberFormat = "#,##0"
        End With

        With .PivotFields("Target Sales")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Total Target"
            .NumberFormat = "#,##0"
        End With

        With .PivotFields("Units Sold")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Sum of Units Sold"
            .NumberFormat = "#,##0"
        End With

        With .PivotFields("Order ID")
            .Orientation = xlDataField
            .Function = xlCount
            .Name = "Total Orders"
            .NumberFormat = "#,##0"
        End With

        On Error Resume Next
        .CalculatedFields.Add "Achievement %", "='Sales Amount'/'Target Sales'", True
        On Error GoTo 0

        On Error Resume Next
        With .PivotFields("Achievement %")
            .Orientation = xlDataField
            .Function = xlSum
            .Name = "Achievement %"
            .NumberFormat = "0.00%"
        End With
        On Error GoTo 0

        On Error Resume Next
        .DataPivotField.Orientation = xlRowField
        .DataPivotField.Position = 1
        On Error GoTo 0

        On Error Resume Next
        .PivotFields("Total Sales").Position = 1
        .PivotFields("Total Target").Position = 2
        .PivotFields("Achievement %").Position = 3
        .PivotFields("Total Orders").Position = 4
        .PivotFields("Sum of Units Sold").Position = 5
        On Error GoTo 0

        .RowAxisLayout xlTabularRow
        .DisplayFieldCaptions = False
        .ShowTableStyleRowStripes = False
        .ColumnGrand = False
        .RowGrand = False
    End With
End Sub

Private Sub CreateKPICards(ws As Worksheet, wsPivot As Worksheet)

    Const KPI_ROW As Long = 200
    ws.Cells(KPI_ROW, 1).Formula = "=Pivot!L5"
    ws.Cells(KPI_ROW, 2).Formula = "=Pivot!L6"
    ws.Cells(KPI_ROW, 3).Formula = "=Pivot!L7"
    ws.Cells(KPI_ROW, 4).Formula = "=Pivot!L8"
    ws.Cells(KPI_ROW, 5).Formula = "=Pivot!L9"
    ws.Rows(KPI_ROW).Hidden = True

    Dim kpiTitles(4) As String
    kpiTitles(0) = "Total Sales"
    kpiTitles(1) = "Total Target"
    kpiTitles(2) = "Total Orders"
    kpiTitles(3) = "Achievement %"
    kpiTitles(4) = "Units Sold"

    Dim kpiCells(4) As String
    kpiCells(0) = "=Pivot!L5"
    kpiCells(1) = "=Pivot!L6"
    kpiCells(2) = "=Pivot!L7"
    kpiCells(3) = "=Pivot!L8"
    kpiCells(4) = "=Pivot!L9"

    Dim accentColors(4) As Long
    accentColors(0) = RGB(37,  99, 235)
    accentColors(1) = RGB(5,  150, 105)
    accentColors(2) = RGB(245, 158,  11)
    accentColors(3) = RGB(124,  58, 237)
    accentColors(4) = RGB(220,  38,  38)

    Dim areaLeft  As Single: areaLeft  = ws.Range("A1").Left + 5
    Dim areaWidth As Single: areaWidth = ws.Range("A1:R1").Width - 10
    Dim cardGap   As Single: cardGap   = 8
    Dim cardW     As Single: cardW     = Int((areaWidth - 4 * cardGap) / 5)
    Dim cardH     As Single: cardH     = 85
    Dim accentH   As Single: accentH   = 12
    Dim cardTop   As Single: cardTop   = ws.Range("A4").Top + 4

    Dim k As Integer
    For k = 0 To 4
        Dim cLeft As Single
        cLeft = areaLeft + k * (cardW + cardGap)

        Dim cardShape As Shape
        Set cardShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, cLeft, cardTop, cardW, cardH)
        With cardShape
            .Name                  = "KPI_Card_" & k
            .Fill.ForeColor.RGB    = RGB(255, 255, 255)
            .Fill.Transparency     = 0
            .Line.ForeColor.RGB    = RGB(218, 224, 235)
            .Line.Weight           = 0.75
            .Shadow.Type           = msoShadow21
            .Shadow.Transparency   = 0.75
            .Shadow.OffsetX        = 1
            .Shadow.OffsetY        = 2
            .Shadow.Size           = 100
            .Shadow.Blur           = 4
        End With

        Dim acBar As Shape
        Set acBar = ws.Shapes.AddShape(msoShapeRectangle, cLeft, cardTop, cardW, accentH)
        With acBar
            .Name               = "KPI_Accent_" & k
            .Fill.ForeColor.RGB = accentColors(k)
            .Fill.Transparency  = 0
            .Line.Visible       = msoFalse
        End With

        Dim titleBox As Shape
        Set titleBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            cLeft + 6, cardTop + accentH + 5, cardW - 12, 22)
        With titleBox
            .Name          = "KPI_Title_" & k
            .Line.Visible  = msoFalse
            .Fill.Visible  = msoFalse
            With .TextFrame
                .Characters.Text           = kpiTitles(k)
                .Characters.Font.Name      = "Calibri"
                .Characters.Font.Size      = 20
                .Characters.Font.Bold      = True
                .Characters.Font.Color     = RGB(10, 10, 10)
                .HorizontalAlignment       = xlHAlignCenter
                .VerticalAlignment         = xlVAlignCenter
            End With
        End With

        Dim valTop  As Single: valTop  = cardTop + accentH + 30
        Dim valH    As Single: valH    = cardH - accentH - 33
        Dim valBox  As Shape
        Set valBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            cLeft + 4, valTop, cardW - 8, valH)
        With valBox
            .Name          = "KPI_Value_" & k
            .Line.Visible  = msoFalse
            .Fill.Visible  = msoFalse
            .DrawingObject.Formula = kpiCells(k)
            With .TextFrame
                .Characters.Font.Name      = "Calibri"
                .Characters.Font.Size      = 24
                .Characters.Font.Bold      = True
                .Characters.Font.Color     = RGB(15, 32, 65)
                .HorizontalAlignment       = xlHAlignCenter
                .VerticalAlignment         = xlVAlignCenter
            End With
        End With

    Next k

End Sub

Private Sub UpdateKPIValues(ws As Worksheet, wsPivot As Worksheet)
    ' Nothing to do - shapes are formula-linked to Pivot!L5:L9
End Sub

Private Sub CreateCharts(ws As Worksheet, wsPivot As Worksheet)

    Const CHART_COLOR_BG   As Long = 16777215
    Const CHART_COLOR_PLOT As Long = 16448765
    Const CHART_COLOR_NAVY As Long = 984097
    Const COLOR_BLUE   As Long = 2493939
    Const COLOR_GREEN  As Long = 984966

    Dim startL   As Single: startL   = ws.Range("A1").Left + 5
    Dim contentW As Single: contentW = ws.Range("A1:R1").Width - 10
    Dim cGap     As Single: cGap     = 8

    Dim rw1Top  As Single: rw1Top  = ws.Range("A12").Top + 3
    Dim rw1H    As Single: rw1H    = ws.Range("A12:A26").Height - 6
    Dim rw2Top  As Single: rw2Top  = ws.Range("A27").Top + 3
    Dim rw2H    As Single: rw2H    = ws.Range("A27:A40").Height - 6

    Dim cht1W As Single: cht1W = Int(contentW * 0.37)
    Dim cht2W As Single: cht2W = Int(contentW * 0.28)
    Dim cht3W As Single: cht3W = contentW - cht1W - cht2W - 2 * cGap
    Dim cht4W As Single: cht4W = Int(contentW * 0.27)
    Dim cht5W As Single: cht5W = contentW - cht4W - cGap

    Dim c1L As Single: c1L = startL
    Dim c2L As Single: c2L = c1L + cht1W + cGap
    Dim c3L As Single: c3L = c2L + cht2W + cGap
    Dim c4L As Single: c4L = startL
    Dim c5L As Single: c5L = c4L + cht4W + cGap

    ' Chart 1: Monthly Sales Trend
    Dim cht1 As ChartObject
    Set cht1 = ws.ChartObjects.Add(c1L, rw1Top, cht1W, rw1H)
    cht1.Name = "chtMonthlyTrend"
    With cht1.Chart
        .ChartType = xlLine
        .HasTitle  = True
        .ChartTitle.Text           = "Monthly Sales Trend"
        .ChartTitle.Font.Size      = 11
        .ChartTitle.Font.Bold      = True
        .ChartTitle.Font.Name      = "Calibri"
        .ChartTitle.Font.Color     = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("A2").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend            = True
        .Legend.Position      = xlLegendPositionBottom
        .Legend.Font.Size     = 8
        .Legend.Font.Name     = "Calibri"
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        Dim ser1 As Series
        On Error Resume Next
        Set ser1 = .SeriesCollection(1)
        If Not ser1 Is Nothing Then
            ser1.Format.Line.ForeColor.RGB = RGB(37, 99, 235)
            ser1.Format.Line.Weight        = 2.25
            ser1.MarkerStyle               = xlMarkerStyleCircle
            ser1.MarkerSize                = 4
            ser1.MarkerForegroundColor     = RGB(37, 99, 235)
            ser1.MarkerBackgroundColor     = RGB(255, 255, 255)
            ser1.HasDataLabels             = True
            ser1.DataLabels.NumberFormat   = "#,##0,"" K"""
            ser1.DataLabels.Font.Size      = 7
            ser1.DataLabels.Font.Color     = RGB(37, 99, 235)
            ser1.DataLabels.Position       = xlLabelPositionAbove
        End If
        Dim ser2 As Series
        Set ser2 = .SeriesCollection(2)
        If Not ser2 Is Nothing Then
            ser2.Format.Line.ForeColor.RGB = RGB(220, 38, 38)
            ser2.Format.Line.Weight        = 1.75
            ser2.Format.Line.DashStyle     = msoLineDash
            ser2.MarkerStyle               = xlMarkerStyleNone
            ser2.HasDataLabels             = False
        End If
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    ' Chart 2: Sales by Region
    Dim cht2 As ChartObject
    Set cht2 = ws.ChartObjects.Add(c2L, rw1Top, cht2W, rw1H)
    cht2.Name = "chtRegion"
    With cht2.Chart
        .ChartType = xlColumnClustered
        .HasTitle  = True
        .ChartTitle.Text       = "Sales by Region"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Name  = "Calibri"
        .ChartTitle.Font.Color = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("A20").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend = False
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(37, 99, 235)
            .GapWidth                  = 60
            .HasDataLabels             = True
            .DataLabels.NumberFormat   = "#,##0,"" K"""
            .DataLabels.Font.Size      = 8
            .DataLabels.Font.Bold      = True
            .DataLabels.Font.Color     = RGB(15, 32, 65)
            .DataLabels.Position       = xlLabelPositionOutsideEnd
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    ' Chart 3: Sales by Manager
    Dim cht3 As ChartObject
    Set cht3 = ws.ChartObjects.Add(c3L, rw1Top, cht3W, rw1H)
    cht3.Name = "chtManager"
    With cht3.Chart
        .ChartType = xlBarClustered
        .HasTitle  = True
        .ChartTitle.Text       = "Sales by Manager"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Name  = "Calibri"
        .ChartTitle.Font.Color = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("A35").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend = False
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(5, 150, 105)
            .GapWidth                  = 40
            .HasDataLabels             = True
            .DataLabels.NumberFormat   = "#,##0,"" K"""
            .DataLabels.Font.Size      = 8
            .DataLabels.Font.Bold      = True
            .DataLabels.Font.Color     = RGB(15, 32, 65)
            .DataLabels.Position       = xlLabelPositionOutsideEnd
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    ' Chart 4: Category Contribution (Donut)
    On Error Resume Next
    ws.ChartObjects("chtCategory").Delete
    On Error GoTo 0

    Dim cht4 As ChartObject
    Set cht4 = ws.ChartObjects.Add(c4L, rw2Top, cht4W, rw2H)
    cht4.Name = "chtCategory"

    With cht4.Chart
        .ChartType = xlDoughnut
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("F2").PivotTable.TableRange1
        On Error GoTo 0
        .HasTitle                   = True
        .ChartTitle.Text            = "Category Contribution"
        .ChartTitle.Font.Size       = 11
        .ChartTitle.Font.Bold       = True
        .ChartTitle.Font.Name       = "Calibri"
        .ChartTitle.Font.Color      = CHART_COLOR_NAVY
        .HasLegend                  = True
        .Legend.Position            = xlLegendPositionRight
        .Legend.Font.Size           = 9
        .Legend.Font.Name           = "Calibri"
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        .PlotArea.Interior.Color    = CHART_COLOR_BG
        On Error Resume Next
        .SeriesCollection(1).DoughnutHoleSize = 40
        Dim dColors(3) As Long
        dColors(0) = RGB(37,  99, 235)
        dColors(1) = RGB(5,  150, 105)
        dColors(2) = RGB(245, 158,  11)
        dColors(3) = RGB(220,  38,  38)
        Dim p As Integer
        For p = 1 To .SeriesCollection(1).Points.Count
            If p <= 4 Then _
                .SeriesCollection(1).Points(p).Format.Fill.ForeColor.RGB = dColors(p - 1)
        Next p
        With .SeriesCollection(1)
            .HasDataLabels               = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.ShowPercentage   = True
            .DataLabels.ShowValue        = False
            .DataLabels.NumberFormat     = "0%"
            .DataLabels.Font.Size        = 8
            .DataLabels.Font.Bold        = True
            .DataLabels.Font.Color       = RGB(255, 255, 255)
        End With
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

    Call CenterDonutPlotArea(cht4)

    ' Chart 5: Top 10 Products by Sales
    Dim cht5 As ChartObject
    Set cht5 = ws.ChartObjects.Add(c5L, rw2Top, cht5W, rw2H)
    cht5.Name = "chtTopProducts"
    With cht5.Chart
        .ChartType = xlColumnClustered
        .HasTitle  = True
        .ChartTitle.Text       = "Top 10 Products by Sales"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Name  = "Calibri"
        .ChartTitle.Font.Color = CHART_COLOR_NAVY
        On Error Resume Next
        .SetSourceData Source:=wsPivot.Range("F35").PivotTable.TableRange1
        On Error GoTo 0
        .HasLegend = False
        .PlotArea.Interior.Color    = RGB(250, 252, 255)
        .ChartArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color   = CHART_COLOR_BG
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 234, 242)
        On Error GoTo 0
        On Error Resume Next
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(245, 158, 11)
            .GapWidth                  = 55
            .HasDataLabels             = True
            .DataLabels.NumberFormat   = "#,##0,"" K"""
            .DataLabels.Font.Size      = 7
            .DataLabels.Font.Bold      = True
            .DataLabels.Font.Color     = RGB(180, 110, 0)
            .DataLabels.Position       = xlLabelPositionOutsideEnd
        End With
        On Error GoTo 0
        On Error Resume Next
        .Axes(xlCategory).TickLabelPosition      = xlTickLabelPositionLow
        .Axes(xlCategory).TickLabels.Orientation = 45
        On Error GoTo 0
        On Error Resume Next: .ShowAllFieldButtons = False: On Error GoTo 0
    End With

End Sub

Private Sub CreateSlicers(wb As Workbook, ws As Worksheet, wsPivot As Worksheet)

    Dim panelLeft   As Single: panelLeft   = ws.Range("T4").Left
    Dim panelTop    As Single: panelTop    = ws.Range("T4").Top
    Dim panelWidth  As Single: panelWidth  = ws.Range("T4:V4").Width
    Dim panelHeight As Single: panelHeight = ws.Range("T4:V40").Height

    Dim panelBG As Shape
    Set panelBG = ws.Shapes.AddShape(msoShapeRectangle, _
        panelLeft, panelTop, panelWidth, panelHeight)
    With panelBG
        .Name               = "SlicerPanel_BG"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency  = 0
        .Line.ForeColor.RGB = RGB(210, 218, 232)
        .Line.Weight        = 0.75
        .ZOrder msoSendToBack
    End With

    Dim slicerW   As Single: slicerW   = panelWidth - 8
    Dim slicerGap As Single: slicerGap = 8

    Dim slicerH(2) As Single
    slicerH(0) =  90
    slicerH(1) = 120
    slicerH(2) = 250

    Dim slicerCols(2) As Integer
    slicerCols(0) = 2
    slicerCols(1) = 1
    slicerCols(2) = 2

    Dim slicerLeft     As Single: slicerLeft     = panelLeft + 4
    Dim slicerTopStart As Single: slicerTopStart = panelTop  + 4

    Dim slicerFields(2) As String
    slicerFields(0) = "Region"
    slicerFields(1) = "Product Category"
    slicerFields(2) = "Month"

    Dim slicerNames(2) As String
    slicerNames(0) = "slcRegion"
    slicerNames(1) = "slcCategory"
    slicerNames(2) = "slcMonth"

    Dim ptNames(6) As String
    ptNames(0) = "pvtMonthly"
    ptNames(1) = "pvtRegion"
    ptNames(2) = "pvtManager"
    ptNames(3) = "pvtCategory"
    ptNames(4) = "pvtSegment"
    ptNames(5) = "pvtTopProducts"
    ptNames(6) = "PT_KPI"

    Dim existSC As SlicerCache
    Dim scToDelete() As String
    Dim scCount As Integer
    scCount = 0

    For Each existSC In wb.SlicerCaches
        ReDim Preserve scToDelete(scCount)
        scToDelete(scCount) = existSC.Name
        scCount = scCount + 1
    Next existSC

    Dim d As Integer
    For d = 0 To scCount - 1
        On Error Resume Next
        wb.SlicerCaches(scToDelete(d)).Delete
        On Error GoTo 0
    Next d

    Dim f As Integer
    Dim currentTop As Single
    currentTop = slicerTopStart

    For f = 0 To 2
        Dim sc As SlicerCache
        Dim sl As Slicer
        Dim firstPT As PivotTable

        On Error Resume Next
        Set firstPT = wsPivot.PivotTables("pvtMonthly")

        Set sc = wb.SlicerCaches.Add2(firstPT, slicerFields(f), slicerNames(f))

        Dim pn As Integer
        For pn = 1 To 6
            On Error Resume Next
            sc.PivotTables.AddPivotTable wsPivot.PivotTables(ptNames(pn))
            On Error GoTo 0
        Next pn

        Set sl = sc.Slicers.Add(ws, , slicerNames(f) & "_visual", _
            slicerFields(f), currentTop, slicerLeft, slicerW, slicerH(f))

        With sl
            .Style           = "SlicerStyleDark4"
            .Left            = slicerLeft
            .Top             = currentTop
            .Width           = slicerW
            .Height          = slicerH(f)
            .NumberOfColumns = slicerCols(f)
            .RowHeight       = 20
        End With

        On Error GoTo 0

        currentTop = currentTop + slicerH(f) + slicerGap
    Next f

End Sub

Private Sub CenterDonutPlotArea(chtObj As ChartObject)
    On Error Resume Next
    Dim cht  As Chart:  Set cht  = chtObj.Chart
    Dim caW  As Single: caW  = cht.ChartArea.Width
    Dim caH  As Single: caH  = cht.ChartArea.Height

    With cht.PlotArea
        .Left   = 0
        .Top    = 0
        .Width  = caW
        .Height = caH
    End With
    With cht.PlotArea
        .Left   = 0
        .Top    = 0
        .Width  = caW
        .Height = caH
    End With
    On Error GoTo 0
End Sub

Private Sub FixDonutCenter()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    If ws Is Nothing Then Exit Sub
    Call CenterDonutPlotArea(ws.ChartObjects("chtCategory"))
    On Error GoTo 0
End Sub
