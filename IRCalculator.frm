VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "IRRIGATION REQUIREMENT CALCULATOR"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5280
   OleObjectBlob   =   "IRCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

    Dim a As Long, LastRow7 As Long, ws7 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    LastRow7 = ws7.Range("A" & Rows.Count).End(xlUp).Row
    
    For a = 2 To LastRow7
        Me.ComboBox1.AddItem ws7.Cells(a, "A").Value
    Next a

End Sub

Private Sub ComboBox1_Change()

    Dim a As Long, LastRow7 As Long, ws7 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    LastRow7 = ws7.Range("A" & Rows.Count).End(xlUp).Row
    
    For a = 2 To LastRow7
        If Me.ComboBox1.Value = ws7.Cells(a, "A").Value Then
            Me.TextBox1.Value = ws7.Cells(a, "B").Value
            Me.TextBox2.Value = ws7.Cells(a, "C").Value
            Me.TextBox3.Value = ws7.Cells(a, "D").Value
        End If
    Next a

End Sub

Private Sub CommandButton1_Click()

    Dim a As Long, LastRow7 As Long, ws7 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    LastRow7 = ws7.Range("A" & Rows.Count).End(xlUp).Row
    
        
    For a = 2 To LastRow7
        If Me.ComboBox1.Value = ws7.Cells(a, "A") Then
            ws7.Cells(a, "B").Value = Me.TextBox1.Value
            ws7.Cells(a, "B").Interior.Color = vbCyan
            ws7.Cells(a, "C").Value = Me.TextBox2.Value
            ws7.Cells(a, "C").Interior.Color = vbCyan
            ws7.Cells(a, "D").Value = Me.TextBox3.Value
            ws7.Cells(a, "D").Interior.Color = vbCyan
        End If
    Next a
    
End Sub

Private Sub CommandButton2_Click()
    Dim a As Long, LastRow7 As Long, ws7 As Worksheet, ws21 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    Set ws21 = Sheets("Final Report Sheet")
    LastRow7 = ws7.Range("A" & Rows.Count).End(xlUp).Row
        If Me.TextBox1.Value = "" And Me.TextBox2.Value = "" And Me.TextBox3.Value = "" Then
            For a = 2 To LastRow7
                ws7.Cells(a, "B").Value = "No Input"
                ws7.Cells(a, "B").Interior.Color = vbMagenta
                ws7.Cells(a, "C").Value = 0
                ws7.Cells(a, "C").Interior.Color = vbMagenta
                ws7.Cells(a, "D").Value = 0
                ws7.Cells(a, "D").Interior.Color = vbMagenta
            Next a
        End If
     
    MsgBox "Assessment Done", vbInformation, "NOTE!"
    ws7.Calculate
    MsgBox "Total Irrigation Water Demand is: " & ws7.Cells("32", "E").Value & " Cubic Metres Per Day", vbInformation, "NOTE!"
    
    Me.Hide
    UserForm1.Hide
    ws7.Visible = xlSheetVisible
    ws7.Activate
    
    UserForm1.TextBox6.Value = ws7.Cells("32", "E").Value
    ws21.Cells("35", "B").Value = ws7.Cells("32", "E").Value
    ws21.Cells("35", "B").Interior.Color = vbCyan
    UserForm1.TextBox6.Enabled = False

End Sub
