Attribute VB_Name = "Module8"
Sub Macro12()
Attribute Macro12.VB_ProcData.VB_Invoke_Func = " \n14"
    Columns("B:B").Select
    Range("B2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub
Sub Macro13()
Attribute Macro13.VB_ProcData.VB_Invoke_Func = " \n14"
    Columns("B:B").Select
    Range("B2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "CURRENT DAM"
    Range("B5").Select
    Columns("B:B").EntireColumn.AutoFit
End Sub
