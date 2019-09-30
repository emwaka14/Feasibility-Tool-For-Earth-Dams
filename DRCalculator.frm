VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "DOMESTIC REQUIREMENT CALCULATOR"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "DRCalculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim b As Long, LastRow As Long, ws8 As Worksheet
    Set ws8 = Sheets("Domestic Water Sheet")
End Sub

Private Sub CommandButton1_Click()
    Dim b As Long, LastRow As Long, ws8 As Worksheet, ws21 As Worksheet
    Set ws8 = Sheets("Domestic Water Sheet")
    Set ws21 = Sheets("Final Report Sheet")
    If Me.TextBox18.Value = "" Or Me.TextBox19.Value = "" Then
        MsgBox ("Please Fill in the details needed")
    Else
        ws8.Cells("1", "B") = Me.TextBox18.Value
        ws8.Cells("1", "B").Interior.Color = vbCyan
        ws8.Cells("2", "B") = Me.TextBox19.Value
        ws8.Cells("2", "B").Interior.Color = vbCyan
         
        MsgBox "Assessment Done!"
        ws8.Calculate
        MsgBox "Total Domestic Water Demand is: " & ws8.Cells("3", "B").Value & " Cubic Metres Per Day"
    
        Me.Hide
        UserForm1.Hide
        ws8.Visible = xlSheetVisible
        ws8.Activate
    
        UserForm1.TextBox28.Value = ws8.Cells("3", "B").Value
        ws21.Cells("33", "B").Value = ws8.Cells("3", "B").Value
        ws21.Cells("33", "B").Interior.Color = vbCyan
        UserForm1.TextBox28.Enabled = False
    End If

End Sub
