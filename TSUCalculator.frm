VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "TSU CALCULATOR"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5460
   OleObjectBlob   =   "TSUCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    Dim k As Long, LastRow3 As Long, ws6 As Worksheet, ws21 As Worksheet
    Set ws6 = Sheets("Livestock Water Sheet")
    Set ws21 = Sheets("Final Report Sheet")
    LastRow3 = ws6.Range("A" & Rows.Count).End(xlUp).Row
    
    Me.CheckBox1.Value = False
    Me.CheckBox2.Value = False
    Me.CheckBox3.Value = False
    Me.CheckBox4.Value = False
    Me.CheckBox5.Value = False
    Me.CheckBox6.Value = False
    Me.CheckBox7.Value = False
    Me.CheckBox8.Value = False
    Me.TextBox1.Visible = False
    Me.TextBox2.Visible = False
    Me.TextBox3.Visible = False
    Me.TextBox4.Visible = False
    Me.TextBox5.Visible = False
    Me.TextBox6.Visible = False
    Me.TextBox7.Visible = False
    Me.TextBox8.Visible = False
End Sub

Private Sub CheckBox1_Click()

    If Me.CheckBox1.Value = True Then
        Me.TextBox1.Visible = True
    Else
        Me.TextBox1.Visible = False
        Me.TextBox1.Value = 0
    End If
    
End Sub

Private Sub CheckBox3_Click()
    
    If Me.CheckBox3.Value = True Then
        Me.TextBox8.Visible = True
    Else
        Me.TextBox8.Visible = False
        Me.TextBox8.Value = 0
    End If

End Sub

Private Sub CheckBox2_Click()

    If Me.CheckBox2.Value = True Then
        Me.TextBox7.Visible = True
    Else
        Me.TextBox7.Visible = False
        Me.TextBox7.Value = 0
    End If
    
End Sub
    
Private Sub CheckBox5_Click()
    
    If Me.CheckBox5.Value = True Then
        Me.TextBox6.Visible = True
    Else
        Me.TextBox6.Visible = False
        Me.TextBox6.Value = 0
    End If

End Sub

Private Sub CheckBox4_Click()

    If Me.CheckBox4.Value = True Then
        Me.TextBox2.Visible = True
    Else
        Me.TextBox2.Visible = False
        Me.TextBox2.Value = 0
    End If
        
End Sub

Private Sub CheckBox6_Click()


    If Me.CheckBox6.Value = True Then
        Me.TextBox3.Visible = True
    Else
        Me.TextBox3.Visible = False
        Me.TextBox3.Value = 0
    End If

End Sub

Private Sub CheckBox7_Click()

    If Me.CheckBox7.Value = True Then
        Me.TextBox4.Visible = True
    Else
        Me.TextBox4.Visible = False
        Me.TextBox4.Value = 0
    End If
    
End Sub

Private Sub CheckBox8_Click()

    If Me.CheckBox8.Value = True Then
        Me.TextBox5.Visible = True
    Else
        Me.TextBox5.Visible = False
        Me.TextBox5.Value = 0
    End If
    
End Sub

Private Sub CommandButton9_Click()

    Dim k As Long, LastRow3 As Long, ws6 As Worksheet, ws21 As Worksheet
    Set ws6 = Sheets("Livestock Water Sheet")
    Set ws21 = Sheets("Final Report Sheet")
    LastRow3 = ws6.Range("A" & Rows.Count).End(xlUp).Row
    
    ws6.Cells("5", "C").Value = Me.TextBox1.Value
    ws6.Cells("9", "C").Value = Me.TextBox2.Value
    ws6.Cells("10", "C").Value = Me.TextBox3.Value
    ws6.Cells("3", "C").Value = Me.TextBox4.Value
    ws6.Cells("7", "C").Value = Me.TextBox5.Value
    ws6.Cells("8", "C").Value = Me.TextBox6.Value
    ws6.Cells("6", "C").Value = Me.TextBox7.Value
    ws6.Cells("4", "C").Value = Me.TextBox8.Value
    
    For k = 3 To LastRow3
    
        If ws6.Cells(k, "C").Value = "" Then
            ws6.Cells(k, "C").Interior.Color = vbMagenta
        Else
            ws6.Cells(k, "C").Interior.Color = vbCyan
        End If
    
    Next k
    
    MsgBox "Assessment Done!"
    ws6.Calculate
    MsgBox "Total Livestock Water Demand is: " & ws6.Cells("11", "D").Value & " Cubic Metres Per Day"
    
    Me.Hide
    UserForm1.Hide
    ws6.Visible = xlSheetVisible
    ws6.Activate
    
    UserForm1.TextBox20.Value = ws6.Cells("11", "D").Value
    UserForm1.TextBox20.Enabled = False
    ws6.Cells("16", "C").Value = UserForm1.TextBox20.Value
    ws6.Cells("14", "C").Clear
    ws21.Cells("34", "B").Value = UserForm1.TextBox20.Value
    ws21.Cells("34", "B").Interior.Color = vbCyan
    
End Sub
