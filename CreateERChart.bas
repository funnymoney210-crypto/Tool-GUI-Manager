Attribute VB_Name = "CreateERChart"






Sub image_toChart()


Dim CurrentChart As Chart
Dim FName As String


ChartName = "ER_Chart"

FName = ThisWorkbook.Path & "\temp.gif"

Set CurrentChart = ThisWorkbook.Sheets("SAT.calc").ChartObjects(ChartName).Chart
CurrentChart.Export fileName:=FName, FilterName:="GIF"

UserForm1.imgChart.Picture = LoadPicture(FName)



End Sub












Sub CheckAndDeleteERChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartExists As Boolean
    Dim shp As Shape

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SAT.calc") ' Change "SAT.calc" to your sheet name if different

    ' Check if the chart "ER_Chart" exists
    chartExists = False
    For Each chartObj In ws.ChartObjects
        If chartObj.Name = "ER_Chart" Then
            chartExists = True
            Exit For
        End If
    Next chartObj

    ' Delete the chart if it exists
    If chartExists Then
        chartObj.Delete
    End If

    ' Delete LCL and UCL lines if they exist
    For Each shp In ws.Shapes
        If shp.Type = msoLine Then
            If shp.Line.ForeColor.RGB = RGB(255, 0, 0) Or shp.Line.ForeColor.RGB = RGB(255, 0, 0) Then
                shp.Delete
            End If
        End If
    Next shp
End Sub



















Sub CreateERTrendChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim rng As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SAT.calc") ' Change "Sheet1" to your sheet name

    ' Find the last row with data in column N
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row


    ' Define the range of data
    Set rng = ws.Range("N1:O" & lastRow)  ' Adjust the range according to your data ' Adjust the range according to your data

    ' Create a new chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    With chartObj
        .Name = "ER_Chart"
        With .Chart
            .SetSourceData Source:=rng
            .ChartType = xlLineMarkers
            .HasTitle = True
            .ChartTitle.Text = "Trend of ER"
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Date"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "ER [u/min]"
            .Axes(xlCategory).CategoryNames = ws.Range("N2:N" & lastRow) ' Adjust the range according to your data
            .HasLegend = False



            ' Get the Y-values of the series
            yValues = .SeriesCollection(1).Values

            ' Color points based on ER value
            For i = 1 To UBound(yValues)
                If yValues(i) > 1.2 Or yValues(i) < 1 Then
                    .SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Red color
                End If
            Next i
            

        End With
    End With



 ' Get chart dimensions and Y-axis scale
    chartTop = chartObj.Top
    chartHeight = chartObj.Height
    yAxisMin = chartObj.Chart.Axes(xlValue).MinimumScale
    yAxisMax = chartObj.Chart.Axes(xlValue).MaximumScale

    ' Calculate the position for LCL and UCL lines
    Dim lclPosition As Double
    Dim uclPosition As Double
    lclPosition = chartTop + chartHeight * (1 - (1 - yAxisMin) / (yAxisMax - yAxisMin))
    uclPosition = chartTop + chartHeight * (1 - (1.2 - yAxisMin) / (yAxisMax - yAxisMin))

    ' Add LCL line as shape
    Set lclShape = ws.Shapes.AddLine(chartObj.Left, lclPosition, chartObj.Left + chartObj.Width, lclPosition)
    With lclShape.Line
        .ForeColor.RGB = RGB(255, 0, 0) ' Blue color
        .Weight = 2
    End With

    ' Add UCL line as shape
    Set uclShape = ws.Shapes.AddLine(chartObj.Left, uclPosition, chartObj.Left + chartObj.Width, uclPosition)
    With uclShape.Line
        .ForeColor.RGB = RGB(255, 0, 0) ' Red color
        .Weight = 2
    End With



End Sub

