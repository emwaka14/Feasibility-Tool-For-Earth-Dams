VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainProgram 
   Caption         =   "Buil-DAM"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3015
   OleObjectBlob   =   "MainProgram.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim j As Worksheet, k As Worksheet, l As Worksheet, a As Worksheet, b As Worksheet, c As Worksheet, d As Worksheet, e As Worksheet, f As Worksheet, g As Worksheet, h As Worksheet, i As Worksheet
    Set a = Sheets("Water Quality Sheet")
    Set b = Sheets("Geotechnical Sheet")
    Set c = Sheets("Domestic Water Sheet")
    Set d = Sheets("Livestock Water Sheet")
    Set e = Sheets("Irrigation Water Sheet")
    Set f = Sheets("Hydrological Analysis Sheet")
    Set g = Sheets("Storage Requirement Sheet")
    Set h = Sheets("HVA Table Sheet")
    Set i = Sheets("Final Report Sheet")
    Set j = Sheets("Final Embankment")
    Set k = Sheets("Cost Estimate Sheet")
    Set l = Sheets("Geotechnical Sheet 2")
    
    a.Visible = xlSheetVeryHidden
    b.Visible = xlSheetVeryHidden
    c.Visible = xlSheetVeryHidden
    d.Visible = xlSheetVeryHidden
    e.Visible = xlSheetVeryHidden
    f.Visible = xlSheetVeryHidden
    g.Visible = xlSheetVeryHidden
    h.Visible = xlSheetVeryHidden
    i.Visible = xlSheetVisible
    i.Activate
    j.Visible = xlSheetVeryHidden
    k.Visible = xlSheetVeryHidden
    l.Visible = xlSheetVeryHidden
End Sub
Private Sub CommandButton1_Click()
    Dim g As Long, LastRow21 As Long, ws21 As Worksheet, Ask1 As Integer
    Set ws21 = Sheets("Final Report Sheet")
    LastRow21 = ws21.Range("B" & Rows.Count).End(xlUp).Row
    
    ws21.Activate
    ws21.Columns("B:B").Select
    ws21.Range("B2").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws21.Range("B2").Select
    ActiveCell.FormulaR1C1 = "CURRENT ASSESSMENT"
    ws21.Range("C2").Select
    ActiveCell.FormulaR1C1 = "PREVIOUS ASSESSMENT"
    ws21.Columns("B:B").EntireColumn.AutoFit
    ws21.Columns("C:C").EntireColumn.AutoFit
            
    Ask1 = MsgBox("If you proceed from here, you may overwrie existing data. Click Yes to overwrite the data or click No to start a new report.", vbYesNo + vbQuestion, "CLEAR DATA")
    
    If Ask1 = vbYes Then
            Module1.ClearDomestic
            Module1.ClearHVA
            Module1.ClearIrrigation
            Module1.ClearLiveStock
            Module1.ClearPrep
            Module1.ClearSoil
            Module1.ClearStorageReq
            Module1.ClearToxicity
            Module1.ClearCostEst
        Me.Hide
        SingleSite.Show
        
    Else
        MsgBox "Note that all tables will be overwritten except for the final report!", vbExclamation, "Warning!"
        
        Me.Hide
        SingleSite.Show
    End If
End Sub

Private Sub CommandButton2_Click()
MultiSite.Show
MainProgram.Hide
End Sub
