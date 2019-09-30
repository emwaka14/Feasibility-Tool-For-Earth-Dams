VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SingleSite 
   Caption         =   "Single Site Analysis"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   OleObjectBlob   =   "SingleSite.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SingleSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox6_Click()
    If Me.CheckBox6.Value = True Then
        Me.CommandButton2.Enabled = True
    ElseIf Me.CheckBox6.Value = False And Me.CheckBox9.Value = False And Me.CheckBox7.Value = False Then
        Me.CommandButton2.Enabled = False
    End If
End Sub

Private Sub CheckBox7_Click()
    If Me.CheckBox7.Value = True Then
        Me.CommandButton2.Enabled = True
    ElseIf Me.CheckBox6.Value = False And Me.CheckBox9.Value = False And Me.CheckBox7.Value = False Then
        Me.CommandButton2.Enabled = False
    End If
End Sub

Private Sub CheckBox9_Click()
    If Me.CheckBox9.Value = True Then
        Me.CommandButton2.Enabled = True
    ElseIf Me.CheckBox6.Value = False And Me.CheckBox9.Value = False And Me.CheckBox7.Value = False Then
        Me.CommandButton2.Enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()
    If Me.CheckBox6.Value = False And Me.CheckBox9.Value = False And Me.CheckBox7.Value = False Then
        Me.CommandButton2.Enabled = False
    Else: Me.CommandButton2.Enabled = True
    End If
End Sub

Private Sub CommandButton2_Click()
    Dim g As Long, LastRow21 As Long, ws21 As Worksheet
    Set ws21 = Sheets("Final Report Sheet")
    
    ws21.Cells("3", "B").Value = Me.TextBox2.Value
    ws21.Cells("3", "B").Interior.Color = vbCyan
    ws21.Cells("4", "B").Value = Me.TextBox3.Value
    ws21.Cells("4", "B").Interior.Color = vbCyan
    ws21.Cells("5", "B").Value = Me.TextBox4.Value
    ws21.Cells("5", "B").Interior.Color = vbCyan
    ws21.Cells("6", "B").Value = Me.TextBox5.Value
    ws21.Cells("6", "B").Interior.Color = vbCyan
    ws21.Cells("7", "B").Value = Me.TextBox6.Value
    ws21.Cells("7", "B").Interior.Color = vbCyan
    
    SingleSite.Hide
    UserForm1.MultiPage1.Value = 0
    UserForm1.Show
End Sub


