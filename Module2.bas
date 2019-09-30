Attribute VB_Name = "Module2"

Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("I11").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = "=""Cumulative Storage Volume"""
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet5!$F$2:$F$14"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""Flooded Area"""
    ActiveChart.FullSeriesCollection(2).Values = "=Sheet5!$D$2:$D$14"
    ActiveChart.FullSeriesCollection(2).XValues = "=Sheet5!$C$2:$C$14"
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Depth"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Depth"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Delete
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveSheet.ChartObjects("Chart 15").Activate
    ActiveChart.Legend.Select
    Selection.Left = 257.454
    Selection.Top = 79.249
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 15").IncrementLeft 276.25
    ActiveSheet.Shapes("Chart 15").IncrementTop 192.5
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveSheet.Shapes("Chart 22").IncrementLeft 147.5
    ActiveSheet.Shapes("Chart 22").IncrementTop 63.75
    ActiveSheet.Shapes("Chart 22").ScaleWidth 1.8194444444, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 22").ScaleHeight 1.4924770341, msoFalse, _
        msoScaleFromTopLeft
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = "=Sheet5!$F$1"
    ActiveChart.FullSeriesCollection(1).Values = "=Sheet5!$F$2:$F$14"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=Sheet5!$D$1"
    ActiveChart.FullSeriesCollection(2).Values = "=Sheet5!$D$2:$D$14"
    ActiveChart.FullSeriesCollection(2).XValues = "=Sheet5!$C$2:$C$14"
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    Selection.Formula = ""
    ActiveChart.ChartTitle.Text = "HVA GRAPH"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "HVA GRAPH"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "DEPTH (m)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "DEPTH (m)"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 9).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementLegendRight)
End Sub
