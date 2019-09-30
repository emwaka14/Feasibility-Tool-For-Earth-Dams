Attribute VB_Name = "Module4"
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("C11:D11").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16776960
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("E11:F11").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.ChartObjects("Chart 10").Activate
    ActiveChart.Parent.Delete
End Sub
