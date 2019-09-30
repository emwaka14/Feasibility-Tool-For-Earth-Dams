Attribute VB_Name = "Module1"
Sub Project1()
MainProgram.Show
End Sub

Sub WaterQual0()
UserForm1.MultiPage1.Value = 1
UserForm1.Show
End Sub

Sub WaterQual1()
UserForm1.MultiPage1.Value = 2
UserForm1.Show
End Sub

Sub WaterQual2()
UserForm1.MultiPage1.Value = 3
UserForm1.Show
End Sub

Sub WaterQual3()
UserForm1.MultiPage1.Value = 4
UserForm1.Show
End Sub

Sub WaterQual4()
Dim ws20 As Worksheet
Set ws20 = Sheets("Storage Requirement Sheet")
    If ws20.Cells("7", "B").Value = "" Then
        UserForm1.MultiPage1.Value = 3
        UserForm1.Show
    Else
        UserForm1.MultiPage1.Value = 5
        UserForm1.Show
    End If
End Sub

Sub WaterQual5()
UserForm1.MultiPage1.Value = 6
UserForm1.Show
End Sub

Sub WaterQual6()
UserForm1.MultiPage1.Value = 7
UserForm1.Show
End Sub

Sub WaterQual7()
UserForm1.MultiPage1.Value = 8
UserForm1.Show
End Sub

Sub FinalREport()
    Dim ws6 As Worksheet
    Set ws6 = Sheets("Final Report Sheet")
    ws6.Visible = xlSheetVisible
    ws6.Activate
End Sub


Sub ClearLiveStock()
Dim k As Long, LastRow3 As Long, ws6 As Worksheet
    Set ws6 = Sheets("Livestock Water Sheet")
    LastRow3 = ws6.Range("A" & Rows.Count).End(xlUp).Row
    
    For k = 3 To LastRow3
        ws6.Cells(k, "C").Clear
    Next k
        ws6.Cells("16", "C").Value = 0
    ws6.Calculate
End Sub

Sub ClearIrrigation()
Dim a As Long, LastRow7 As Long, ws7 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    LastRow7 = ws7.Range("A" & Rows.Count).End(xlUp).Row
    
    For a = 2 To LastRow7
        ws7.Cells(a, "B").Clear
        ws7.Cells(a, "C").Clear
        ws7.Cells(a, "C").Value = 0
        ws7.Cells(a, "D").Clear
        ws7.Cells(a, "D").Value = 0
    Next a
    ws7.Calculate
End Sub

Sub ClearDomestic()
    Dim b As Long, LastRow8 As Long, ws8 As Worksheet
    Set ws8 = Sheets("Domestic Water Sheet")
    LastRow8 = ws8.Range("A" & Rows.Count).End(xlUp).Row
    
    For b = 1 To LastRow8 - 1
        ws8.Cells(b, "B").Clear
        ws8.Cells(b, "B").Value = 0
    Next b
    ws8.Calculate
End Sub

Sub ClearToxicity()
    
    Dim i As Long, LastRow As Long, ws As Worksheet
    Set ws = Sheets("Water Quality Sheet")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 3 To LastRow
        ws.Cells(i, "H").Clear
        ws.Cells(i, "H").Value = 0
        ws.Cells(i, "I").Clear
        ws.Cells(i, "J").Clear
        ws.Cells(i, "K").Clear
    Next i
        
End Sub

Sub ClearHVA()
    Dim c As Long, LastRow9 As Long, ws9 As Worksheet
    Set ws9 = Sheets("HVA Table Sheet")
    LastRow9 = ws9.Range("B" & Rows.Count).End(xlUp).Row
    
    For c = 2 To LastRow9
        ws9.Cells(c, "B").Clear
        ws9.Cells(c, "C").Clear
        ws9.Cells(c, "D").Clear
        ws9.Cells(c, "E").Clear
        ws9.Cells(c, "F").Clear
    Next c
    ws9.Calculate
End Sub

Sub ClearSoil()
Dim d As Long, ws10 As Worksheet, LastRow As Long
    Set ws10 = Sheets("Geotechnical Sheet 2")
    LastRow = ws10.Range("A" & Rows.Count).End(xlUp).Row
    ws10.Range("M3:M5").ClearContents
    ws10.Range("M3:M5").Interior.Color = vbWhite
    
    For d = 1 To LastRow
        ws10.Range("A" & d & ":J" & d).Interior.Color = vbWhite
    Next d
End Sub

Sub ClearPrep()
Dim e As Long, LastRow11 As Long, ws11 As Worksheet
    Set ws11 = Sheets("Hydrological Analysis Sheet")
    ws11.Range("B2:B4").Clear
    ws11.Cells("7", "B").Clear
    ws11.Calculate

End Sub

Sub ClearStorageReq()
    Dim f As Long, LastRo20 As Long, ws20 As Worksheet
    Set ws20 = Sheets("Storage Requirement Sheet")
    
        ws20.Range("B2:B4").Clear
        ws20.Range("B7:B9").Clear
        ws20.Calculate
End Sub

Sub ClearCostEst()
    Dim f As Long, LastRo20 As Long, ws20 As Worksheet
    Set ws20 = Sheets("Cost Estimate Sheet")
    
        ws20.Range("B2:B4").Clear
        ws20.Calculate
End Sub

Sub ClearOptimumStorage()
    Dim h As Long, z As Long, LastRow50 As Long, ws9 As Worksheet
    Set ws9 = Sheets("HVA Table Sheet")
    LastRow50 = ws9.Range("C" & Rows.Count).End(xlUp).Row
    Dim ws6 As Worksheet
    Set ws6 = Sheets("Livestock Water Sheet")
    Dim ws7 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    Dim ws8 As Worksheet
    Set ws8 = Sheets("Domestic Water Sheet")
    Dim ws20 As Worksheet
    Set ws20 = Sheets("Storage Requirement Sheet")
    Dim ws11 As Worksheet
    Set ws11 = Sheets("Hydrological Analysis Sheet")
    
    For h = 2 To LastRow50
        If ws9.Range("C" & h).Interior.Color = vbGreen Then
            ws9.Range("C" & h & ":D" & h).Interior.Color = vbCyan
            ws9.Range("E" & h & ":F" & h).Interior.Color = vbWhite
        End If
    Next h
    
    ws9.ChartObjects.Delete
End Sub

