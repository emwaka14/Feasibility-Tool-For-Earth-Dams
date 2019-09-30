Attribute VB_Name = "Module3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.CutCopyMode = False
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=Sheet5!$AZ$2:$AZ$14"
    ActiveChart.FullSeriesCollection(3).Name = "=""Corresponding Flooded Area"""
    ActiveChart.FullSeriesCollection(3).Values = "=Sheet5!$AZ$2:$AZ$14"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).Name = "=""Chosen Volume"""
    ActiveChart.FullSeriesCollection(4).Name = "=""Corresponding Volume"""
    ActiveChart.FullSeriesCollection(4).Values = "=Sheet5!$AY$2:$AY$14"
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMinorGridlines = True
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
End Sub
