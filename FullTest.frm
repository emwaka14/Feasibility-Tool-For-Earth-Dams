VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ANALYSIS WIZARD"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520.001
   OleObjectBlob   =   "FullTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox13_Change()

    Dim i As Long, LastRow As Long, wsa As Worksheet
    Set wsa = Sheets("Geotechnical Sheet 2")
    LastRow = wsa.Range("J" & Rows.Count).End(xlUp).Row
    
    For i = 3 To LastRow
        If Me.ComboBox13.Value = wsa.Cells(i, "J").Value Then
            Me.TextBox41.Value = wsa.Cells(i, "K").Value
        End If
    Next i
    
End Sub

Private Sub ComboBox15_Change()

    Dim i As Long, LastRow As Long, wsa As Worksheet
    Set wsa = Sheets("Geotechnical Sheet")
    LastRow = wsa.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 5 To LastRow
        If Me.ComboBox15.Value = wsa.Cells(i, "A").Value Then
            Me.TextBox47.Value = wsa.Cells(i, "F").Value
        End If
    Next i
    
End Sub

Private Sub ComboBox18_Change()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    If (Me.ComboBox18.Value = "Blanket" Or Me.ComboBox18.Value = "Chimney") And Me.TextBox51.Value <= 7.5 Then
        MsgBox "Recommended drain for that embankment height is the Toe Drain.", vbExclamation, "NOTE!"
        Me.ComboBox18.Value = "Toe"
    ElseIf (Me.ComboBox18.Value = "Toe") And Me.TextBox51.Value > 7.5 Then
        MsgBox "Recommended drain for that embankment height is the Blanket Drain or Chimney Drain.", vbExclamation, "NOTE!"
        Me.ComboBox18.Value = "Blanket"
    End If
    frs.Range("B57").Value = Me.ComboBox18.Value
    frs.Cells("57", "B").Interior.Color = vbCyan
End Sub

Private Sub CommandButton42_Click()

    Dim i As Long, LastRow As Long, wsa As Worksheet
    Set wsa = Sheets("Geotechnical Sheet")
    LastRow = wsa.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 5 To LastRow
        If Me.ComboBox15.Value = wsa.Cells(i, "A").Value Then
            wsa.Cells(i, "F").Value = Me.TextBox47.Value
            wsa.Cells(i, "F").Interior.Color = vbCyan
        End If
    Next i
    
End Sub

Private Sub CommandButton38_Click()

    Dim i As Long, LastRow As Long, wsa As Worksheet
    Set wsa = Sheets("Geotechnical Sheet 2")
    LastRow = wsa.Range("J" & Rows.Count).End(xlUp).Row
    
    For i = 3 To LastRow
        If Me.ComboBox13.Value = wsa.Cells(i, "L").Value Then
            wsa.Cells(i, "M").Value = Me.TextBox41.Value
            wsa.Cells(i, "M").Interior.Color = vbCyan
        End If
    Next i
    
End Sub

Private Sub CommandButton53_Click()

    Dim frs As Worksheet, ws11 As Worksheet
    Set frs = Sheets("Final Report Sheet")
    Set ws11 = Sheets("Hydrological Analysis Sheet")
    
    ws11.Cells("2", "B").Value = Me.TextBox38.Value
    frs.Cells("8", "B").Value = Me.TextBox38.Value
    frs.Cells("8", "B").Interior.Color = vbCyan
        
    If Me.TextBox35.Value > 500 Then
        MsgBox "Fields are too far from the proposed dam site!", vbExclamation, "Qn.1 WARNING!"
    Else
        MsgBox "Site distance from fields is OK!", vbInformation, "Qn.1 NOTE!"
    End If
    frs.Range("B11").Value = Me.TextBox35.Value
    frs.Cells("11", "B").Interior.Color = vbCyan

    If Me.OptionButton36.Value = True Then
        MsgBox "Conveying water without gravitational help is not cost effective", vbExclamation, "Qn.2 WARNING!"
        frs.Range("B12").Value = "FALSE"
        frs.Range("B12").Interior.Color = vbRed
    Else
        MsgBox "Site elevation from field is OK!", vbInformation, "Qn.1 NOTE!"
        frs.Range("B12").Value = "TRUE"
        frs.Range("B12").Interior.Color = vbGreen
    End If

    If Me.TextBox36.Value > 5000 Then
        MsgBox "Proposed dam site too far from human settlement for domestic water access but is optimum for livestock water access.", vbExclamation, "Qn.3 WARNING!"
    ElseIf Me.TextBox36.Value > 10000 Then
        MsgBox "Proposed dam site too far from human settlement for domestic and livestock water access!", vbExclamation, "Qn.3 WARNING!"
    Else:
        MsgBox "Proposed dam site is situated at a distance optimum for both domestic and livestock consumption.", vbInformation, "Qn.3 NOTE!"
    End If
    frs.Range("B13").Value = Me.TextBox36.Value
    frs.Cells("13", "B").Interior.Color = vbCyan

    If Me.OptionButton33.Value = True Then
        If Me.TextBox37.Value > 55 Then
            MsgBox "The valley is too steep, it is likely to affect site access and the stability of the reservior", vbExclamation, "Qn.4 WARNING!"
        ElseIf Me.TextBox37.Value < 55 And Me.TextBox37.Value > 3 Then
            MsgBox "The site slope is optimum for earth dam construction.", vbInformation, "Qn.4 NOTE!"
        ElseIf Me.OptionButton34.Value = True Then
            MsgBox "It is advisable to get a gentle sloping valley. Flat land do not have the best storage.", vbExclamation, "Qn.5 WARNING!"
        End If
    End If
    frs.Range("B14").Value = Me.TextBox37.Value
    frs.Cells("14", "B").Interior.Color = vbCyan
    
    If Me.OptionButton30.Value = True Then
        MsgBox "The lack of gulleys to the proposed reservoir site limits its capability to fill up. Not a good site", vbExclamation, "Qn.7 WARNING!"
        frs.Range("B15").Value = "FALSE"
        frs.Range("B15").Interior.Color = vbRed
    ElseIf Me.OptionButton29.Value = True Then
        MsgBox "The presence of gulleys to the proposed reservoir site increases its capability to fill up. It is a good site.", vbInformation, "Qn.7 NOTE!"
        frs.Range("B15").Value = "TRUE"
        frs.Range("B15").Interior.Color = vbGreen
    End If

    If Me.OptionButton37.Value = True Then
        MsgBox "The site is satisfactorily accessible by all machinery", vbInformation, "Qn.8 NOTE!"
        frs.Range("B16").Value = "ALL MACHINERY"
        frs.Cells("16", "B").Interior.Color = vbCyan
    ElseIf Me.OptionButton38.Value = True Then
        MsgBox "Site excavations and works a limited to small machinery.", vbExclamation, "Qn.8 WARNING!"
        frs.Range("B16").Value = "SMALL MACHINERY"
        frs.Cells("16", "B").Interior.Color = vbCyan
    ElseIf Me.OptionButton39.Value = True Then
        MsgBox "Site excavations are very limited! Mass excavations will not be possible!", vbExclamation, "Qn.2 WARNING!"
        frs.Range("B16").Value = "ON-FOOT"
        frs.Cells("16", "B").Interior.Color = vbCyan
    End If
    Me.CommandButton36.Enabled = True

End Sub

Private Sub CommandButton36_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")

    If Me.TextBox35.Value <= 500 And Me.OptionButton35.Value = True And Me.TextBox36.Value <= 5000 And Me.TextBox37.Value < 55 And Me.OptionButton29.Value = True Then
        MsgBox "SITE LOCATION IS FEASIBLE!", vbInformation, "NOTE!"
        Me.MultiPage1.Value = 1
        frs.Range("B10").Value = "PASS"
        frs.Range("B10").Interior.Color = vbGreen
    Else
        Dim Ask1 As Integer
        Ask1 = MsgBox("SITE LOCATION IS NOT FEASIBLE! PROCEED INSPITE OF THIS?", vbQuestion + vbYesNo, "QUESTION!")
        If Ask1 = vbYes Then
            Me.MultiPage1.Value = 1
        Else
            MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
            frs.Visible = xlSheetVisible
            frs.Activate
            Me.Hide
        End If
        frs.Range("B10").Value = "FAIL"
        frs.Range("B10").Interior.Color = vbRed
    End If
    
End Sub

Private Sub CommandButton41_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    If Me.TextBox42.Value <= 3 Then
        MsgBox "The site seepage rate is good!", vbInformation, "SITE INFILTRATION TEST: NOTE"
            Me.Label69.Enabled = False
            Me.CommandButton52.Enabled = False
        Me.Frame36.Enabled = False
            Me.Label72.Enabled = False
            Me.OptionButton44.Enabled = False
            Me.OptionButton45.Enabled = False
        Me.Frame48.Enabled = False
            Me.Label93.Enabled = False
            Me.OptionButton56.Enabled = False
            Me.OptionButton55.Enabled = False
        frs.Range("B26").Value = "PASS"
        frs.Range("B26").Interior.Color = vbGreen

    ElseIf Me.TextBox42.Value > 3 And Me.TextBox42.Value < 30 Then
        frs.Range("B26").Value = "FAIL"
        frs.Range("B26").Interior.Color = vbRed
        MsgBox "The site is doubtful! Seepage rate is high!", vbExclamation, "SITE INFILTRATION TEST: WARNING!"
        MsgBox "Mitigation Measures can be:" & vbNewLine & "1. Lining with LDPE or HDPE on the upstream face of the dam to where the water sits." & vbNewLine & "2. Excavating the site to a depth where there is an impermeable layer." & vbNewLine & "NB: Both measures lead to extra budget costs.", vbInformation, "MITIGATION MEASURES FOR HIGH SEEPAGE RATES"
        Dim Ask1 As Integer
        Ask1 = MsgBox("Are you considering a mitigation method?", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
        
            If Ask1 = vbYes Then
                MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
            Else:
                MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                frs.Visible = xlSheetVisible
                frs.Activate
                Me.Hide
            End If
        Me.Frame36.Enabled = True
        Me.Label72.Enabled = True
        Me.OptionButton44.Enabled = True
        Me.OptionButton45.Enabled = True
    Else:
        frs.Range("B26").Value = "FAIL"
        frs.Range("B26").Interior.Color = vbRed
        MsgBox "The site is too permeable! Seepage rate is high!", vbCritical, "SITE INFILTRATION TEST: WARNING!"
        MsgBox "Mitigation Measures can be:" & vbNewLine & "1. Lining with LDPE or HDPE on the upstream face of the dam to where the water sits." & vbNewLine & "2. Excavating the site to a depth where there is an impermeable layer." & vbNewLine & "NB: Both measures lead to extra budget costs.", vbInformation, "MITIGATION MEASURES FOR HIGH SEEPAGE RATES"
        Dim Ask10 As Integer
        Ask10 = MsgBox("Are you considering a mitigation method?", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
        
            If Ask10 = vbYes Then
                MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
            Else:
                MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                frs.Visible = xlSheetVisible
                frs.Activate
                Me.Hide
            End If
        Me.Frame36.Enabled = True
        Me.Label72.Enabled = True
        Me.OptionButton44.Enabled = True
        Me.OptionButton45.Enabled = True
        Me.Frame36.Enabled = True
    End If
    frs.Range("B27").Value = Me.TextBox42.Value
    frs.Cells("27", "B").Interior.Color = vbCyan
    
End Sub

Private Sub CommandButton48_Click()

Dim frs As Worksheet
Set frs = Sheets("Final Report Sheet")
If Me.TextBox42.Value <= 3 Then
    MsgBox "The site seepage rate is good!", vbInformation, "SITE INFILTRATION TEST: PASSED!"
ElseIf Me.TextBox42.Value > 3 And Me.TextBox42.Value < 30 Then
        MsgBox "The site seepage rate is doubtful! There is need for a mitigation measure.", vbExclamation, "SITE INFILTRATION TEST: WARNING!"
        If Me.OptionButton44.Value = True Then
            MsgBox "Mitigation Measures can be:" & vbNewLine & "1. Lining with LDPE or HDPE on the upstream face of the dam to where the water sits.", vbInformation, "MITIGATION MEASURES FOR HIGH SEEPAGE RATES"
                Dim Ask1 As Integer
                Ask1 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
                    If Ask1 = vbYes Then
                        MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
                        frs.Range("B29").Value = "PASS"
                        frs.Range("B29").Interior.Color = vbGreen
                    Else:
                        MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                        MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                        frs.Range("B29").Value = "FAIL"
                        frs.Range("B29").Interior.Color = vbRed
                        frs.Visible = xlSheetVisible
                        frs.Activate
                        Me.Hide
                    End If
        ElseIf Me.OptionButton45.Value = True Then
                MsgBox "Mitigation Measures can be:" & vbNewLine & "1. Excavating the site to a depth where there is an impermeable layer.", vbInformation, "MITIGATION MEASURES FOR HIGH SEEPAGE RATES"
                    If Me.TextBox44.Value < 3 Then
                        MsgBox "Excavation to specified depth is OK.", vbInformation, "SITE INFILTRATION TEST: NOTE!"
                            Dim Ask2 As Integer
                            Ask2 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
                                If Ask2 = vbYes Then
                                    MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
                                    frs.Range("B29").Value = "PASS"
                                    frs.Range("B29").Interior.Color = vbGreen
                                Else:
                                    MsgBox "The site is too permeable! You are advised to stop the test and drop the site."
                                    MsgBox "You'll shortly be redirected to the Final Report Sheet."
                                    frs.Visible = xlSheetVisible
                                    frs.Activate
                                    Me.Hide
                                    frs.Range("B29").Value = "FAIL"
                                    frs.Range("B29").Interior.Color = vbRed
                                End If
                    ElseIf Me.TextBox44.Value > 3 Then
                            MsgBox "Excavation to specified depth is uneconomical!", vbCritical, "SITE INFILTRATION TEST: WARNING!"
                                Dim Ask3 As Integer
                                Set frs = Sheets("Final Report Sheet")
                                Ask3 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
                                    If Ask3 = vbYes Then
                                        MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
                                        frs.Range("B29").Value = "PASS"
                                        frs.Range("B29").Interior.Color = vbGreen
                                    Else:
                                        MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                                        MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                                        frs.Visible = xlSheetVisible
                                        frs.Activate
                                        Me.Hide
                                    End If
                    End If
        End If

ElseIf Me.TextBox42.Value > 30 Then
    Dim Ask0 As Integer
    MsgBox "The site is too permeable! Seepage rate is high! You are advised NOT TO PROCEED with this site", vbCritical, "SITE INFILTRATION TEST: WARNING!"
    Ask0 = MsgBox("Do you wish to proceed inspite of the warning?", vbQuestion + vbYesNo, "SITE INFILTRATION TEST: WARNING!")
        If Ask0 = vbYes Then
            If Me.OptionButton44.Value = True Then
                MsgBox "Mitigation Measures can be:" & vbNewLine & "1. Lining with LDPE or HDPE on the upstream face of the dam to where the water sits.", vbInformation, "MITIGATION MEASURES FOR HIGH SEEPAGE RATES"
                    Dim Ask4 As Integer
                        Ask4 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
                            If Ask4 = vbYes Then
                                MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
                            Else:
                                MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                                MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                                frs.Visible = xlSheetVisible
                                frs.Activate
                                Me.Hide
                            End If
            ElseIf Me.OptionButton45.Value = True Then
                    MsgBox "Mitigation Measures can be:" & vbNewLine & "1. Excavating the site to a depth where there is an impermeable layer.", vbInformation, "MITIGATION MEASURES FOR HIGH SEEPAGE RATES"
                        If Me.TextBox44.Value < 3 Then
                            MsgBox "Excavation to specified depth is OK.", vbInformation, "SITE INFILTRATION TEST: NOTE!"
                                Dim Ask5 As Integer
                                Ask5 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
                                    If Ask5 = vbYes Then
                                        MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
                                    Else:
                                        MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                                        MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                                        frs.Visible = xlSheetVisible
                                        frs.Activate
                                        Me.Hide
                                    End If
                        ElseIf Me.TextBox44.Value > 3 Then
                                MsgBox "Excavation to specified depth is uneconomical!", vbCritical, "SITE INFILTRATION TEST: WARNING!"
                                    Dim Ask6 As Integer
                                    Ask6 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
                                        If Ask6 = vbYes Then
                                            MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
                                        Else:
                                            MsgBox "TEST ENDS!", vbExclamation, "NOTE"
                                            MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
                                            frs.Visible = xlSheetVisible
                                            frs.Activate
                                            Me.Hide
                                    End If
                        End If
            End If
        Else:
            MsgBox "You will be directed to the final report!", vbInformation, "NOTE!"
            Dim frs1 As Worksheet
            Set frs1 = Sheets("Final Report Sheet")
            frs1.Visible = xlSheetVisible
            frs1.Activate
            Me.Hide
        End If
End If

MsgBox "DONE!", vbInformation, "INFILTRATION ASSESSMENT!"

If Me.MultiPage2.Value = 0 Then

    Dim i As Long, j As Long, k As Long, LastRow As Long, LastRow2 As Long, ws As Worksheet, ws1 As Worksheet
    Set ws = Sheets("Geotechnical Sheet")
    Set frs = Sheets("Final Report Sheet")
    Set ws1 = Sheets("Geotechnical Sheet 2")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    LastRow2 = ws1.Range("A" & Rows.Count).End(xlUp).Row
    
    i = 3
    Do Until (ws1.Cells(i, "B").Value <= ws1.Cells("3", "M").Value And ws1.Cells(i, "C").Value >= ws1.Cells("3", "M").Value) And (ws1.Cells(i, "D").Value <= ws1.Cells("4", "M").Value And ws1.Cells(i, "E").Value >= ws1.Cells("4", "M").Value) And (ws1.Cells(i, "F").Value <= ws1.Cells("5", "M").Value And ws1.Cells(i, "G").Value >= ws1.Cells("5", "M").Value)
        i = i + 1
    Loop
    
    ws1.Range("A" & i & ":J" & i).Interior.Color = vbGreen
    MsgBox "The Soil type found is: " & ws1.Cells(i, "A").Value, vbInformation, "NOTE!"
    frs.Range("B28").Value = ws1.Cells(i, "A").Value
    frs.Cells("28", "B").Interior.Color = vbCyan
    MsgBox "Please check the suitability for the embankment type for that soil type and make a decision basing on the availablle soil type.", vbInformation, "TEXTURAL CLASS RESULTS"
    MsgBox "DONE!", vbInformation, "BORROW SITE TEST"
    Me.Hide
    ws1.Visible = xlSheetVisible
    ws1.Activate

ElseIf Me.MultiPage2.Value = 1 Then
    
    frs.Range("B28").Value = ""
    If Me.CheckBox1.Value = True Then
        MsgBox "Soil group found is: GW and good for the shell of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = "GW"
    ElseIf Me.CheckBox2.Value = True Then
        MsgBox "Soil group found is: GP and good for the shell of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = "GP"
    ElseIf Me.CheckBox3.Value = True Then
        MsgBox "Soil group found is: GM and good for the shell of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = "GM"
    ElseIf Me.CheckBox4.Value = True Then
        MsgBox "Soil group found is: GC and very good for the core of a zoned dam and very good for a Homogeneous dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = "GC"
    End If
    
    If Me.CheckBox5.Value = True Then
        MsgBox "Soil group found is: SW and good for the shell of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", SW"
    ElseIf Me.CheckBox6.Value = True Then
        MsgBox "Soil group found is: SP and fairly good for the shell of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", SP"
    ElseIf Me.CheckBox7.Value = True Then
        MsgBox "Soil group found is: SM and fairly good for the core of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", SM"
    ElseIf Me.CheckBox8.Value = True Then
        MsgBox "Soil group found is: SC and fairly good for the core of a zoned dam and good for a Homogeneous dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", SC"
    End If
    
    If Me.CheckBox9.Value = True Then
        MsgBox "Soil group found is: CH and fairly good for the core of a zoned dam and fairly good for a Homogeneous dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", CH"
    ElseIf Me.CheckBox10.Value = True Then
        MsgBox "Soil group found is: CL and good for the core of a zoned dam and good for a Homogeneous dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", CL"
    End If
    
    If Me.CheckBox11.Value = True Then
        MsgBox "Soil group found is: MH and poor for the core of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", MH"
    ElseIf Me.CheckBox12.Value = True Then
        MsgBox "Soil group found is: ML and poor for the core of a zoned dam.", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", ML"
    End If
    
    If Me.CheckBox13.Value = True Then
        MsgBox "Soil group found is: OH and NOT SUITABLE for any embankment type!", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", OH"
    ElseIf Me.CheckBox14.Value = True Then
        MsgBox "Soil group found is: OL and NOT SUITABLE for any embankment type!", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", OL"
    End If
    
    If Me.CheckBox15.Value = True Then
        MsgBox "Soil group found is: Pt and NOT SUITABLE for any embankment type!", vbInformation, "NOTE!"
        frs.Range("B28").Value = frs.Range("B28").Value & ", Pt"
    End If
    frs.Cells("28", "B").Interior.Color = vbCyan

    MsgBox "Please don't use any soil type categorised as NOT SUITABLE and you're advised to terminate the assessment.", vbExclamation, "WARNING!"
End If
Me.CommandButton9.Enabled = True
End Sub

Private Sub CommandButton49_Click()
Me.MultiPage1.Value = 6
End Sub

Private Sub CommandButton50_Click()
    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    MsgBox "ASSESSMENT COMPLETE! You will be shortly redirected to the Final Report Sheet.", vbInformation, "ASSESSMENT COMPLETE!"
        frs.Visible = xlSheetVisible
        frs.Activate
        frs.Columns("B:B").EntireColumn.AutoFit
        Me.Hide
End Sub

Private Sub CommandButton51_Click()
Me.MultiPage1.Value = 0
End Sub

Private Sub CommandButton52_Click()
    If Me.TextBox44.Value > 3 Then
        MsgBox "Excavation to specified depth is uneconomical!", vbCritical, "SITE INFILTRATION TEST: WARNING!"
        Me.Frame48.Enabled = True
            Me.Label93.Enabled = True
            Me.OptionButton56.Enabled = True
            Me.OptionButton55.Enabled = True
    ElseIf Me.TextBox44.Value < 3 Then
        MsgBox "Excavation to specified depth is OK.", vbInformation, "SITE INFILTRATION TEST: NOTE!"
        Me.Frame48.Enabled = False
            Me.Label93.Enabled = False
            Me.OptionButton56.Enabled = False
            Me.OptionButton55.Enabled = False
    End If
End Sub

Private Sub CommandButton54_Click()
Me.MultiPage1.Value = 8
End Sub

Private Sub CommandButton55_Click()
    Dim ws As Worksheet, frs As Worksheet, x As Double
    Set ws = Sheets("Cost Estimate Sheet")
    Set frs = Sheets("Final Report Sheet")
        If Me.TextBox51.Value < 5 Then
            x = 4.5
        ElseIf Me.TextBox51.Value >= 5 And Me.TextBox51.Value < 10 Then
            x = 5
        ElseIf Me.TextBox51.Value >= 10 And Me.TextBox51.Value <= 15 Then
            x = 5.5
        End If
    ws.Cells("2", "B").Value = Me.TextBox61.Value
    ws.Cells("3", "B").Value = 0.216 * (Me.TextBox51.Value) * (Me.TextBox61.Value) * (2 * Me.TextBox55.Value + x * Me.TextBox51.Value)
    MsgBox "The estimated volume of earthworks is: " & ws.Cells("3", "B").Value & " cubic metres.", vbInformation, "DONE!"
    frs.Range("B67").Value = ws.Cells("2", "B").Value
    frs.Cells("67", "B").Interior.Color = vbCyan
    frs.Range("B68").Value = ws.Cells("3", "B").Value
    frs.Cells("68", "B").Interior.Color = vbCyan
    
    ws.Visible = xlSheetVisible
    ws.Activate
    Me.Hide
End Sub

Private Sub CommandButton56_Click()
    Dim ws As Worksheet, frs As Worksheet
    Set ws = Sheets("Cost Estimate Sheet")
    Set frs = Sheets("Final Report Sheet")
    ws.Cells("4", "B").Value = Me.TextBox62.Value
    ws.Calculate
    MsgBox "The estimated cost of earthmoving is $" & ws.Cells("5", "B").Value & ".", vbInformation, "DONE!"
    frs.Range("B69").Value = ws.Cells("5", "B").Value
    frs.Cells("69", "B").Interior.Color = vbCyan
    ws.Visible = xlSheetVisible
    ws.Activate
    Me.Hide
End Sub

Private Sub CommandButton57_Click()
Me.MultiPage1.Value = 7
End Sub

Private Sub CommandButton58_Click()

    Dim frs As Worksheet, calc As Worksheet
    Set frs = Sheets("Final Report Sheet")
    Set calc = Sheets("CALCULATOR")
    
    frs.Range("B70").Value = Me.TextBox63.Value
    frs.Range("B70").Interior.Color = vbCyan
    calc.Range("B1").Value = Me.TextBox63.Value
    calc.Calculate
    Me.TextBox52.Value = calc.Range("B3").Value
    calc.Calculate
    Me.TextBox53.Value = calc.Range("B4").Value
    calc.Calculate
    Me.TextBox51.Value = calc.Range("B5").Value
    calc.Calculate
    Me.TextBox55.Value = calc.Range("B6").Value
    
    frs.Range("B55").Value = calc.Range("B4").Value
    frs.Cells("55", "B").Interior.Color = vbCyan
    frs.Range("B56").Value = calc.Range("B5").Value
    frs.Cells("56", "B").Interior.Color = vbCyan
    frs.Range("B58").Value = calc.Range("B6").Value
    frs.Cells("58", "B").Interior.Color = vbCyan
    frs.Range("B59").Value = calc.Range("B3").Value
    frs.Cells("59", "B").Interior.Color = vbCyan
    
    If Me.TextBox51.Value < 5 Then
        Me.TextBox56.Value = "2.5:1"
        Me.TextBox60.Value = "2.0:1"
        frs.Range("B60").Value = "2.5:1"
        frs.Cells("60", "B").Interior.Color = vbCyan
        frs.Range("B61").Value = "2.0:1"
        frs.Cells("61", "B").Interior.Color = vbCyan
    ElseIf Me.TextBox51.Value >= 5 And Me.TextBox51.Value < 10 Then
        Me.TextBox56.Value = "2.5:1"
        Me.TextBox60.Value = "2.5:1"
        frs.Range("B60").Value = "2.5:1"
        frs.Cells("60", "B").Interior.Color = vbCyan
        frs.Range("B61").Value = "2.5:1"
        frs.Cells("61", "B").Interior.Color = vbCyan
    ElseIf Me.TextBox51.Value >= 10 And Me.TextBox51.Value <= 15 Then
        Me.TextBox56.Value = "3.0:1"
        Me.TextBox60.Value = "2.5:1"
        frs.Range("B60").Value = "3.0:1"
        frs.Cells("60", "B").Interior.Color = vbCyan
        frs.Range("B61").Value = "2.5:1"
        frs.Cells("61", "B").Interior.Color = vbCyan
    End If
    
    If Me.TextBox51.Value <= 7.5 Then
        Me.ComboBox18.Value = "Toe"
    ElseIf Me.TextBox51.Value <= 15 And Me.TextBox51.Value > 7.5 Then
        Me.ComboBox18.Value = "Blanket"
    End If
    
    If Me.OptionButton50.Value = 1 Or Me.OptionButton48.Value = 0 Then
        Me.Label87.Enabled = True
        Me.Label88.Enabled = True
        Me.Label89.Enabled = True
        Me.TextBox57.Value = "1.5:1"
        Me.TextBox58.Value = 1.5 * Me.TextBox51.Value
        Me.TextBox59.Value = 0.75 * Me.TextBox51.Value
        
    ElseIf Me.OptionButton50.Value = 0 Or Me.OptionButton48.Value = 1 Then
        Me.Label87.Enabled = False
        Me.Label88.Enabled = False
        Me.Label89.Enabled = False
        Me.TextBox57.Value = ""
        Me.TextBox58.Value = ""
        Me.TextBox59.Value = ""
    End If
    Me.CommandButton9.Enabled = True
    frs.Range("B62").Value = Me.TextBox57.Value
    frs.Cells("62", "B").Interior.Color = vbCyan
    frs.Range("B63").Value = Me.TextBox59.Value
    frs.Cells("63", "B").Interior.Color = vbCyan
    frs.Range("B64").Value = Me.TextBox58.Value
    frs.Cells("64", "B").Interior.Color = vbCyan
    
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub MultiPage2_Change()
If Me.MultiPage2.Value = 0 Then
    MsgBox "You are advised to reject any soils containing PEAT or any ORGANIC MATTER!", vbExclamation, "TEXTURAL CLASS USAGE WARNING!"
Else
End If
End Sub

Private Sub MultiPage2_Click(ByVal Index As Long)
If Index = 0 Then
    MsgBox "You are advised to reject any soils containing PEAT or any ORGANIC MATTER!", vbExclamation, "TEXTURAL CLASS USAGE WARNING!"
End If
End Sub

Private Sub OptionButton13_Click()
    Me.CommandButton1.Enabled = True
End Sub

Private Sub OptionButton14_Click()
    If Me.OptionButton14.Value = True Then
        Me.CommandButton1.Enabled = True
    Else
        Me.CommandButton1.Enabled = False
    End If
    
End Sub

Private Sub OptionButton15_Click()

    If Me.OptionButton15.Value = True Then
        Me.CommandButton1.Enabled = True
        Me.Label51.Visible = False
        Me.Frame15.Visible = False
        Me.Label51.Enabled = False
        Me.Frame15.Enabled = False
    Else:
        Me.CommandButton1.Enabled = True
        Me.Label51.Visible = True
        Me.Frame15.Visible = True
        Me.Label51.Enabled = True
        Me.Frame15.Enabled = True
    End If
End Sub

Private Sub OptionButton16_Click()

    Me.CommandButton1.Enabled = False
    If Me.OptionButton16.Value = True Then
        Me.Label51.Visible = True
        Me.Frame15.Visible = True
        Me.Label51.Enabled = True
        Me.Frame15.Enabled = True
        If Me.OptionButton14.Value = True Then
            Me.CommandButton1.Enabled = True
        ElseIf Me.OptionButton13.Value = True Then
            Me.CommandButton1.Enabled = False
        End If
    Else:
        Me.CommandButton1.Enabled = False
        Me.Label51.Visible = False
        Me.Frame15.Visible = False
        Me.Label51.Enabled = False
        Me.Frame15.Enabled = False
    End If
    
End Sub

Private Sub OptionButton29_Click()

End Sub

Private Sub OptionButton30_Click()

End Sub

Private Sub OptionButton33_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    
    frs.Cells("17", "B").Value = "TRUE"
    frs.Cells("17", "B").Interior.Color = vbGreen
    
    If Me.OptionButton33.Value = True Then
        Me.Label56.Enabled = True
        Me.TextBox37.Enabled = True
    Else
    End If
End Sub

Private Sub OptionButton34_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    
    frs.Cells("17", "B").Value = "FALSE"
    frs.Cells("17", "B").Interior.Color = vbRed
    
    If Me.OptionButton34.Value = True Then
        Me.Label56.Enabled = False
        Me.TextBox37.Enabled = False
    Else
    End If
End Sub

Private Sub OptionButton44_Click()
    If Me.OptionButton44.Value = True Then
            Me.Label69.Enabled = False
            Me.CommandButton52.Enabled = False
            Me.TextBox44.Enabled = False
            Me.Frame48.Enabled = True
            Me.Label93.Enabled = True
            Me.OptionButton55.Enabled = True
            Me.OptionButton56.Enabled = True
    Else
            Me.Label69.Enabled = True
            Me.CommandButton52.Enabled = True
            Me.TextBox44.Enabled = True
            Me.Frame48.Enabled = True
            Me.Label93.Enabled = False
            Me.OptionButton55.Enabled = False
            Me.OptionButton56.Enabled = False
    End If
End Sub

Private Sub OptionButton45_Click()
    If Me.OptionButton45.Value = True Then
            Me.Label69.Enabled = True
            Me.CommandButton52.Enabled = True
            Me.TextBox44.Enabled = True
            Me.Frame48.Enabled = False
            Me.Label93.Enabled = False
            Me.OptionButton55.Enabled = False
            Me.OptionButton56.Enabled = False
    Else
            Me.Label69.Enabled = False
            Me.CommandButton52.Enabled = False
            Me.TextBox44.Enabled = False
    End If
End Sub

Private Sub OptionButton46_Click()
    Me.OptionButton54.Enabled = False
    Me.OptionButton53.Enabled = False
    Me.OptionButton43.Enabled = True
    Me.OptionButton42.Enabled = True
    Me.Label85.Enabled = False
    Me.Label73.Enabled = True
    
    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    MsgBox "Mitigation strategy needed!", vbExclamation, "ROCK FOUNDATION ASSESSMENT: WARNING!"
    MsgBox "Mitigation Measures can be:" & vbNewLine & "1. For cracks; They can be sealed with cement (Quite expensive)" & vbNewLine & "2. For Weathered rocks and searms, there is no mitigation measure. The site is NOT GOOD!", vbInformation, "MITIGATION MEASURES FOR ROCK FOUNDATIONS"
    Ask2 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
        If Ask2 = vbYes Then
            MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
            frs.Range("B29").Value = "PASS"
            frs.Range("B29").Interior.Color = vbRed
        Else:
            MsgBox "The foundation is not good!", vbExclamation, "CLAY FOUNDATION ASSESSMENT: WARNING!"
            MsgBox "You'll shortly be redirected to the Final Report Sheet.", vbInformation, "CLAY FOUNDATION ASSESSMENT: NOTE!"
            frs.Visible = xlSheetVisible
            frs.Activate
            Me.Hide
            frs.Range("B29").Value = "FAIL"
            frs.Range("B29").Interior.Color = vbRed
        End If
End Sub

Private Sub OptionButton47_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    Me.OptionButton54.Enabled = False
    Me.OptionButton53.Enabled = False
    Me.OptionButton43.Enabled = False
    Me.OptionButton42.Enabled = False
    Me.Label85.Enabled = False
    Me.Label73.Enabled = False
    MsgBox "Foundation is good!", vbInformmation, "ROCK FOUNDATION ASSESSMENT: NOTE!"
    frs.Range("B29").Value = "PASS"
    frs.Range("B29").Interior.Color = vbGreen
End Sub

Private Sub OptionButton48_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    If Me.OptionButton48.Value = True Then
        Me.Label87.Enabled = False
        Me.Label88.Enabled = False
        Me.Label89.Enabled = False
        Me.TextBox57.Enabled = False
        Me.TextBox58.Enabled = False
        Me.TextBox59.Enabled = False
        Me.TextBox57.Value = ""
        Me.TextBox58.Value = ""
        Me.TextBox59.Value = ""
    Else:
        Me.Label87.Enabled = True
        Me.Label88.Enabled = True
        Me.Label89.Enabled = True
    End If
    frs.Range("B65").Value = "HOMOGENEOUS EMBANKMENT"
    frs.Cells("65", "B").Interior.Color = vbCyan
End Sub

Private Sub OptionButton50_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    If Me.OptionButton50.Value = True Then
        Me.Label87.Enabled = True
        Me.Label88.Enabled = True
        Me.Label89.Enabled = True
    Else:
        Me.Label87.Enabled = False
        Me.Label88.Enabled = False
        Me.Label89.Enabled = False
        Me.TextBox57.Enabled = False
        Me.TextBox58.Enabled = False
        Me.TextBox59.Enabled = False
        Me.TextBox57.Value = ""
        Me.TextBox58.Value = ""
        Me.TextBox59.Value = ""
    End If
    frs.Range("B65").Value = "ZONED EMBANKMENT"
    frs.Cells("65", "B").Interior.Color = vbCyan
End Sub

Private Sub OptionButton52_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    Me.OptionButton54.Enabled = False
    Me.OptionButton53.Enabled = False
    Me.OptionButton46.Enabled = True
    Me.OptionButton47.Enabled = True
    Me.OptionButton43.Enabled = False
    Me.OptionButton42.Enabled = False
    Me.Label85.Enabled = False
    Me.Label78.Enabled = True
    Me.Label73.Enabled = False
    frs.Range("B30").Value = "ROCK FOUNDATION"
    frs.Cells("30", "B").Interior.Color = vbCyan
End Sub
Private Sub OptionButton51_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    Me.OptionButton54.Enabled = True
    Me.OptionButton53.Enabled = True
    Me.OptionButton46.Enabled = False
    Me.OptionButton47.Enabled = False
    Me.OptionButton43.Enabled = False
    Me.OptionButton42.Enabled = False
    Me.Label85.Enabled = True
    Me.Label78.Enabled = False
    Me.Label73.Enabled = False
    frs.Range("B30").Value = "CLAY FOUNDATION"
    frs.Cells("30", "B").Interior.Color = vbCyan
End Sub

Private Sub OptionButton53_Click()
    Me.OptionButton46.Enabled = False
    Me.OptionButton47.Enabled = False
    Me.OptionButton43.Enabled = True
    Me.OptionButton42.Enabled = True
    Me.Label78.Enabled = False
    Me.Label73.Enabled = True
    
    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    MsgBox "Mitigation strategy needed!", vbExclamation, "CLAY FOUNDATION ASSESSMENT: WARNING!"
    Ask2 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
        If Ask2 = vbYes Then
            MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
            frs.Range("B29").Value = "PASS"
            frs.Range("B29").Interior.Color = vbGreen
        Else:
            MsgBox "The foundation is not good!", vbExclamation, "CLAY FOUNDATION ASSESSMENT: WARNING!"
            MsgBox "You'll shortly be redirected to the Final Report Sheet.", vbInformation, "CLAY FOUNDATION ASSESSMENT: NOTE!"
            frs.Visible = xlSheetVisible
            frs.Activate
            Me.Hide
            frs.Range("B29").Value = "FAIL"
            frs.Range("B29").Interior.Color = vbRed
        End If
    
End Sub

Private Sub OptionButton54_Click()

    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    Me.OptionButton46.Enabled = False
    Me.OptionButton47.Enabled = False
    Me.OptionButton43.Enabled = False
    Me.OptionButton42.Enabled = False
    Me.Label78.Enabled = False
    Me.Label73.Enabled = False
    MsgBox "Good Foundation!", vbInformation, "CLAY FOUNDATION ASSESSMENT: NOTE!"
    frs.Range("B29").Value = "PASS"
    frs.Range("B29").Interior.Color = vbGreen
End Sub

Private Sub OptionButton57_Click()
    Me.OptionButton46.Enabled = False
    Me.OptionButton47.Enabled = False
    Me.OptionButton43.Enabled = True
    Me.OptionButton42.Enabled = True
    Me.Label85.Enabled = False
    Me.Label78.Enabled = False
    Me.Label73.Enabled = True
    
    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
    MsgBox "Mitigation strategy needed to reduce seepage losses!", vbExclamation, "SAND/GRAVEL FOUNDATION WARNING!"
    Ask2 = MsgBox("Are you considering the (or any other) mitigation method? Note that all mitigation methods involve extra costs!", vbQuestion + vbYesNo, "SITE INFILTRATION TESTS: MITIGATION METHODS")
        If Ask2 = vbYes Then
            MsgBox "TEST PROCEEDS!", vbInformation, "NOTE"
            frs.Range("B29").Value = "PASS"
            frs.Range("B29").Interior.Color = vbGreen
        Else:
            MsgBox "The foundation is not good!", vbExclamation, "SAND/GRAVEL FOUNDATION ASSESSMENT: WARNING!"
            MsgBox "You'll shortly be redirected to the Final Report Sheet.", vbInformation, "CLAY FOUNDATION ASSESSMENT: NOTE!"
            frs.Visible = xlSheetVisible
            frs.Activate
            Me.Hide
            frs.Range("B29").Value = "FAIL"
            frs.Range("B29").Interior.Color = vbRed
        End If
    frs.Range("B30").Value = "SAND/GRAVEL FOUNDATION"
    frs.Cells("30", "B").Interior.Color = vbCyan
End Sub

Private Sub TextBox63_Change()
    If Me.TextBox63.Value = "" Then
        Me.CommandButton58.Enabled = False
    Else
        Me.CommandButton58.Enabled = True
    End If
End Sub

Private Sub TextBox8_Change()
    If Me.TextBox8.Value = "" Then
            Me.CommandButton23.Enabled = False
    Else: Me.CommandButton23.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()

        Me.CommandButton9.Enabled = False
    If Me.OptionButton48.Value = True Then
        Me.Label87.Enabled = False
        Me.Label88.Enabled = False
        Me.Label89.Enabled = False
        Me.TextBox57.Enabled = False
        Me.TextBox58.Enabled = False
        Me.TextBox59.Enabled = False
        Me.TextBox57.Value = ""
        Me.TextBox58.Value = ""
        Me.TextBox59.Value = ""
    Else:
        Me.Label87.Enabled = True
        Me.Label88.Enabled = True
        Me.Label89.Enabled = True
    End If
    
            If Me.TextBox63.Value = "" Then
                Me.CommandButton58.Enabled = False
            Else
                Me.CommandButton58.Enabled = True
            End If


            Me.TextBox51.Enabled = False
            Me.TextBox52.Enabled = False
            Me.TextBox53.Enabled = False
            Me.TextBox54.Enabled = False
            Me.TextBox55.Enabled = False
            Me.TextBox56.Enabled = False
            Me.TextBox57.Enabled = False
            Me.TextBox58.Enabled = False
            Me.TextBox59.Enabled = False
            Me.TextBox60.Enabled = False

            Me.CommandButton36.Enabled = False
    
            Me.Label69.Enabled = False
            Me.CommandButton52.Enabled = False
        Me.Frame36.Enabled = False
            Me.Label72.Enabled = False
            Me.OptionButton44.Enabled = False
            Me.OptionButton45.Enabled = False
        Me.Frame48.Enabled = False
            Me.Label93.Enabled = False
            Me.TextBox44.Enabled = False
            Me.OptionButton56.Enabled = False
            Me.OptionButton55.Enabled = False
            Me.OptionButton54.Enabled = False
            Me.OptionButton53.Enabled = False
            Me.OptionButton46.Enabled = False
            Me.OptionButton47.Enabled = False
            Me.OptionButton43.Enabled = False
            Me.OptionButton42.Enabled = False
            Me.Label85.Enabled = False
            Me.Label78.Enabled = False
            Me.Label73.Enabled = False
    If Me.OptionButton34.Value = True Then
        Me.Label56.Enabled = False
        Me.TextBox37.Enabled = False
    ElseIf Me.OptionButton33.Value = True Then
        Me.Label56.Enabled = True
        Me.TextBox37.Enabled = True
    End If

    Dim i As Long, j As Long, LastRow As Long, LastRow2 As Long, LastRow3 As Long, LastRow10 As Long, wsl As Worksheet, ws As Worksheet, ws1 As Worksheet, wsa As Worksheet, wsb As Worksheet
    Set ws = Sheets("Water Quality Sheet")
    Set wsb = Sheets("Geotechnical Sheet")
    Set wsa = Sheets("Geotechnical Sheet 2")
    Dim c As Long, LastRow9 As Long, Last9 As Long, ws9 As Worksheet
    Set ws9 = Sheets("HVA Table Sheet")
    Set wsl = Sheets("Final Embankment")
    Dim e As Long, LastRow11 As Long, ws11 As Worksheet
    Set ws11 = Sheets("Hydrological Analysis Sheet")
    LastRow11 = ws11.Range("A" & Rows.Count).End(xlUp).Row
    LastRow9 = ws9.Range("A" & Rows.Count).End(xlUp).Row
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    LastRow2 = wsa.Range("L" & Rows.Count).End(xlUp).Row
    LastRow3 = wsb.Range("A" & Rows.Count).End(xlUp).Row
    LastRow10 = wsl.Range("AG" & Rows.Count).End(xlUp).Row
    
    Me.Label51.Visible = False
    Me.Frame15.Visible = False
    Me.CommandButton1.Enabled = False
    Me.CommandButton9.Enabled = False
    Me.CommandButton12.Enabled = False
    Me.CommandButton14.Enabled = False
    Me.CommandButton32.Enabled = False
    Me.CommandButton33.Enabled = False
    Me.CommandButton35.Enabled = False
    
    If Me.TextBox8.Value = "" Then
            Me.CommandButton23.Enabled = False
        Else: Me.CommandButton23.Enabled = False
    End If
    
    
    If SingleSite.CheckBox7.Value = 0 Then
        Me.Frame13.Enabled = False
    End If
    
    If SingleSite.CheckBox9.Value = 0 Then
        Me.Frame12.Enabled = False
    End If
    
    If SingleSite.CheckBox6.Value = 0 Then
        Me.Frame8.Enabled = False
    End If
    
    For i = 3 To LastRow
        Me.ComboBox1.AddItem ws.Cells(i, "A").Value
    Next i
    
    For i = 3 To LastRow2
        Me.ComboBox13.AddItem wsa.Cells(i, "L").Value
    Next i
    
    For i = 5 To LastRow3
        Me.ComboBox15.AddItem wsb.Cells(i, "A").Value
    Next i
    
    For c = 2 To LastRow9
        Me.ComboBox11.AddItem ws9.Cells(c, "A").Value
    Next c
    
     For i = 2 To LastRow10
        Me.ComboBox18.AddItem wsl.Cells(i, "AG").Value
    Next i
    
End Sub

Private Sub ComboBox1_Change()

    Dim i As Long, LastRow As Long, ws As Worksheet
    Set ws = Sheets("Water Quality Sheet")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 3 To LastRow
        If Me.ComboBox1.Value = ws.Cells(i, "A").Value Then
            Me.TextBox1.Value = ws.Cells(i, "H").Value
        End If
    Next i

End Sub

Private Sub CommandButton17_Click()

    Dim i As Long, LastRow As Long, ws As Worksheet
    Set ws = Sheets("Water Quality Sheet")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 3 To LastRow
        If Me.ComboBox1.Value = ws.Cells(i, "A") Then
            ws.Cells(i, "H").Value = Me.TextBox1.Value
            ws.Cells(i, "H").Interior.Color = vbCyan
        End If
    Next i
    

End Sub


Private Sub ComboBox8_Change()

    Dim j As Long, LastRow2 As Long, ws1 As Worksheet
    Set ws1 = Sheets("Geotechnical Sheet")
    LastRow2 = ws1.Range("A" & Rows.Count).End(xlUp).Row
    
    For j = 5 To LastRow2
        If Me.ComboBox8.Value = ws1.Cells(j, "A") Then
            Me.TextBox12.Value = ws1.Cells(j, "F").Value

        End If

    Next j

End Sub


Private Sub CommandButton18_Click()
    Dim j As Long, LastRow2 As Long, ws1 As Worksheet
    Set ws1 = Sheets("Geotechnical Sheet")
    LastRow2 = ws1.Range("A" & Rows.Count).End(xlUp).Row
    
    For j = 5 To LastRow2
        If Me.ComboBox8.Value = ws1.Cells(j, "A") Then
            ws1.Cells(j, "F").Value = Me.TextBox12.Value
            ws1.Cells(j, "F").Interior.Color = vbCyan
    
        End If
    
    Next j
End Sub
Private Sub CommandButton8_Click()

    Dim i As Long, LastRow As Long, ws As Worksheet, ws21 As Worksheet
    Set ws = Sheets("Water Quality Sheet")
    Set ws21 = Sheets("Final Report Sheet")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    If SingleSite.CheckBox7.Value = 0 Then
        MsgBox "Assessment not for Domestic Water", vbInformation, "NOTE!"
        For i = 3 To LastRow
            ws.Cells(i, "I").Clear
            ws.Cells(i, "I").Interior.Color = xlNone
        Next i
    Else:
        For i = 3 To LastRow
            If ws.Cells(i, "C").Value = "" Then
                ws.Cells(i, "I").Value = "Not Needed"
                ws.Cells(i, "I").Interior.Color = vbMagenta
            ElseIf ws.Cells(i, "H").Value <= ws.Cells(i, "C") And ws.Cells(i, "B") <= ws.Cells(i, "H") Then
                ws.Cells(i, "I").Value = "SUITABLE"
                ws.Cells(i, "I").Interior.Color = vbGreen
            Else: ws.Cells(i, "I").Value = "NOT SUITABLE"
                ws.Cells(i, "I").Interior.Color = vbRed
            End If
        Next i
        
        For i = 3 To LastRow
            If ws.Cells(i, "I").Interior.Color = vbRed Then
                ws21.Cells(20, "B").Interior.Color = vbRed
                ws21.Cells(20, "B").Value = "FAIL"
                GoTo End1
            Else: ws21.Cells(20, "B").Interior.Color = vbGreen
                ws21.Cells(20, "B").Value = "PASS"
            End If
        Next i
End1:
    End If
    
    If SingleSite.CheckBox9.Value = 0 Then
        MsgBox "Assessment not for Livestock Water", vbInformation, "NOTE!"
        For i = 3 To LastRow
            ws.Cells(i, "J").Clear
            ws.Cells(i, "J").Interior.Color = xlNone
        Next i
    Else:
        For i = 3 To LastRow
            If ws.Cells(i, "E").Value = "" Then
                ws.Cells(i, "J").Value = "Not Needed"
                ws.Cells(i, "J").Interior.Color = vbMagenta
            ElseIf ws.Cells(i, "H").Value <= ws.Cells(i, "E") And ws.Cells(i, "D") <= ws.Cells(i, "H") Then
                ws.Cells(i, "J").Value = "SUITABLE"
                ws.Cells(i, "J").Interior.Color = vbGreen
            Else: ws.Cells(i, "J").Value = "NOT SUITABLE"
                ws.Cells(i, "J").Interior.Color = vbRed
            End If
        Next i
        
        For i = 3 To LastRow
            If ws.Cells(i, "J").Interior.Color = vbRed Then
                ws21.Cells(21, "B").Interior.Color = vbRed
                ws21.Cells(21, "B").Value = "FAIL"
                GoTo End2
            Else: ws21.Cells(21, "B").Interior.Color = vbGreen
                ws21.Cells(21, "B").Value = "PASS"
            End If
        Next i
End2:
    End If
    
    If SingleSite.CheckBox6.Value = 0 Then
        MsgBox "Assessment not for Irrigation", vbInformation, "NOTE!"
        For i = 3 To LastRow
            ws.Cells(i, "K").Clear
            ws.Cells(i, "K").Interior.Color = xlNone
        Next i
    Else:
        For i = 3 To LastRow
            If ws.Cells(i, "G").Value = "" Then
                ws.Cells(i, "K").Value = "Not Needed"
                ws.Cells(i, "K").Interior.Color = vbMagenta
            ElseIf ws.Cells(i, "H").Value <= ws.Cells(i, "G") And ws.Cells(i, "F") <= ws.Cells(i, "H") Then
                ws.Cells(i, "K").Value = "SUITABLE"
                ws.Cells(i, "K").Interior.Color = vbGreen
            Else: ws.Cells(i, "K").Value = "NOT SUITABLE"
                ws.Cells(i, "K").Interior.Color = vbRed
            End If

        Next i
        
        For i = 3 To LastRow
            If ws.Cells(i, "K").Interior.Color = vbRed Then
                ws21.Cells(22, "B").Interior.Color = vbRed
                ws21.Cells(22, "B").Value = "FAIL"
                GoTo End3
            Else: ws21.Cells(22, "B").Interior.Color = vbGreen
                ws21.Cells(22, "B").Value = "PASS"
            End If
        Next i
End3:
    End If
    Me.CommandButton1.Enabled = True
    MsgBox "ASSESSMENT DONE! Please check worksheet for more information", vbInformation, "NOTE!"
    MsgBox "You are advised not to proceed if the assessment contains MINERAL CONTENT VALUES classified as 'NOT SUITABLE'", vbExclamation, "NOTE!"
    UserForm1.Hide
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

Private Sub CommandButton19_Click()
Me.CommandButton9.Enabled = True
End Sub

Private Sub CommandButton1_Click()

    Dim i As Long, LastRow As Long, ws21 As Worksheet, ws As Worksheet
    Set ws21 = Sheets("Final Report Sheet")
    Set ws = Sheets("Water Quality Sheet")
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
        For i = 3 To LastRow
            If ws.Cells(i, "I").Interior.Color = vbRed Then
                GoTo Exit1
            ElseIf ws.Cells(i, "J").Interior.Color = vbRed Then
                GoTo Exit1
            ElseIf ws.Cells(i, "K").Interior.Color = vbRed Then
                GoTo Exit1
            ElseIf Me.OptionButton13.Value = True Then
                GoTo Exit1
            Else: GoTo Exit2
            End If
        Next i
Exit1: Ask2 = MsgBox("The site is UNSAFE due to possibility of pollution. You are advised NOT TO PROCEED further with the feasibility assessment. Proceed inspite of the state of the site?", vbExclamation + vbYesNo, "SITE UNSAFE!")
    
    If Ask2 = vbYes Then
        Me.MultiPage1.Value = 2
    Else
        MsgBox "You will be directed to the final report the final report!", vbInformation, "NOTE!"
        ws21.Activate
        Me.Hide
    End If
    ws21.Range("B23").Value = "FAIL"
    ws21.Range("B23").Interior.Color = vbRed
Exit2:
    Me.MultiPage1.Value = 2
    If ws21.Range("B23").Value = "" Then
        ws21.Range("B23").Value = "PASS"
        ws21.Range("B23").Interior.Color = vbGreen
    End If
    
End Sub

Private Sub CommandButton10_Click()
Me.MultiPage1.Value = 1
End Sub

Private Sub CommandButton9_Click()
Me.MultiPage1.Value = 3
End Sub

Private Sub CommandButton11_Click()
Me.MultiPage1.Value = 2
End Sub

Private Sub CommandButton12_Click()
Me.MultiPage1.Value = 4
End Sub

Private Sub CommandButton13_Click()
Me.MultiPage1.Value = 3
End Sub

Private Sub CommandButton14_Click()
Me.MultiPage1.Value = 5
End Sub

Private Sub CommandButton31_Click()
Me.MultiPage1.Value = 4
End Sub

Private Sub CommandButton32_Click()
Me.MultiPage1.Value = 6
End Sub

Private Sub CommandButton34_Click()
Me.MultiPage1.Value = 5
End Sub

Private Sub CommandButton35_Click()
Me.MultiPage1.Value = 7
End Sub

Private Sub OptionButton8_Click()
    Me.CommandButton24.Enabled = False
    Me.TextBox20.Enabled = False
    Me.TextBox8.Enabled = True
    If Me.TextBox8.Value = "" Then
        Me.CommandButton23.Enabled = False
    Else: Me.CommandButton23.Enabled = True
    End If
End Sub

Private Sub OptionButton9_Click()
    Me.CommandButton23.Enabled = False
    Me.TextBox8.Enabled = False
    Me.TextBox20.Enabled = False
    Me.CommandButton24.Enabled = True
End Sub


Private Sub CommandButton24_Click()
    UserForm2.Show
End Sub

Private Sub CommandButton21_Click()
    UserForm4.Show
End Sub

Private Sub CommandButton23_Click()
    Dim ws6 As Worksheet, ws21 As Worksheet
    Set ws6 = Sheets("Livestock Water Sheet")
    Set ws21 = Sheets("Final Report Sheet")
        ws6.Cells("14", "C").Value = Me.TextBox8.Value
        Me.TextBox20.Value = 0.05 * Me.TextBox8.Value
        ws6.Cells("16", "C").Value = Me.TextBox20.Value
        ws21.Cells("21", "B").Value = Me.TextBox20.Value
        ws21.Cells("21", "B").Interior.Color = vbCyan
End Sub

Private Sub CommandButton26_Click()
    UserForm5.Show
End Sub


Private Sub ComboBox11_Change()
    Dim c As Long, LastRow9 As Long, Last9 As Long, ws9 As Worksheet
    Set ws9 = Sheets("HVA Table Sheet")
    LastRow9 = ws9.Range("A" & Rows.Count).End(xlUp).Row
    
    For c = 2 To LastRow9
        If Me.ComboBox11.Value = ws9.Cells(c, "A").Value Then
            Me.TextBox34.Value = ws9.Cells(c, "B").Value
           Me.TextBox33.Value = ws9.Cells(c, "C").Value
           Me.TextBox22.Value = ws9.Cells(c, "D").Value
        End If
    Next c
End Sub

Private Sub CommandButton27_Click()
    Dim c As Long, ws9 As Worksheet, LastRow9 As Long
    Set ws9 = Sheets("HVA Table Sheet")
    LastRow9 = ws9.Range("A" & Rows.Count).End(xlUp).Row
    
    For c = 2 To LastRow9
        If Me.ComboBox11.Value = ws9.Cells(c, "A").Value Then
           ws9.Cells(c, "B").Value = Me.TextBox34.Value
           ws9.Cells(c, "B").Interior.Color = vbCyan
           ws9.Cells(c, "C").Value = Me.TextBox33.Value
           ws9.Cells(c, "C").Interior.Color = vbCyan
           ws9.Cells(c, "D").Value = Me.TextBox22.Value
           ws9.Cells(c, "D").Interior.Color = vbCyan
        End If
    Next c
End Sub

Private Sub CommandButton28_Click()

Dim c As Long, d As Long, x As Long, ws9 As Worksheet, LastRow9 As Long
Set ws9 = Sheets("HVA Table Sheet")
LastRow92 = ws9.Range("A" & Rows.Count).End(xlUp).Row
LastRow9 = ws9.Range("C" & Rows.Count).End(xlUp).Row
LastRow91 = ws9.Range("D" & Rows.Count).End(xlUp).Row

For d = 3 To LastRow9
    ws9.Cells(d, "E").Value = ((ws9.Cells(d - 1, "D").Value + ws9.Cells(d, "D").Value) / 2) * (ws9.Cells(d, "B").Value - ws9.Cells(d - 1, "B").Value)
    ws9.Cells(d, "F").Value = ws9.Cells(d - 1, "F").Value + ws9.Cells(d, "E").Value
Next d

ws9.Activate
ws9.Calculate
ws9.Cells("2", "X").Select
For c = 2 To LastRow92
    If ws9.Cells(c, "E").Value < 0 Or ws9.Cells(c, "F").Value < 0 Then
        ws9.Cells(c, "E").Value = ""
        ws9.Cells(c, "F").Value = ""
    End If
Next c
    ws9.Shapes.AddChart2(227, xlLine, 1150, 0, 500, 350).Select
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = ws9.Cells("1", "F").Value
    ActiveChart.FullSeriesCollection(1).Values = ws9.Range("F2:F" & LastRow9)
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = ws9.Cells("1", "D").Value
    ActiveChart.FullSeriesCollection(2).Values = ws9.Range("D2:D" & LastRow9)
    ActiveChart.FullSeriesCollection(2).XValues = ws9.Range("C2:C" & LastRow9)
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
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMinorGridlines = True
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    ActiveChart.ChartArea.Select
    Me.CommandButton33.Enabled = True
    UserForm1.Hide
    ws9.Visible = xlSheetVisible
    ws9.Activate
End Sub

Private Sub CommandButton16_Click()
    Dim e As Long, LastRow11 As Long, ws11 As Worksheet, frs As Worksheet
    Set ws11 = Sheets("Hydrological Analysis Sheet")
    Set frs = Sheets("Final Report Sheet")
    LastRow11 = ws11.Range("A" & Rows.Count).End(xlUp).Row
    
        ws11.Cells("3", "B").Value = Me.TextBox15.Value
        ws11.Cells("4", "B").Value = Me.TextBox21.Value
        ws11.Cells("7", "B").Value = Me.TextBox26.Value
        
        ws11.Calculate
        
        frs.Range("B45").Value = Me.TextBox15.Value
        frs.Range("B45").Interior.Color = vbCyan
        frs.Range("B46").Value = Me.TextBox26.Value
        frs.Range("B46").Interior.Color = vbCyan
        frs.Range("B47").Value = Me.TextBox21.Value
        frs.Range("B47").Interior.Color = vbCyan
        frs.Range("B8").Value = Me.TextBox11.Value
        frs.Range("B8").Interior.Color = vbCyan
        
        MsgBox "Assessment Done!", vbInformation, "NOTE!"
        MsgBox "The Maximum harvestable water for given period is " & ws11.Cells("8", "B").Value & " Cubic Metres.", vbInformation, "NOTE!"
        frs.Range("B48").Value = ws11.Cells("8", "B").Value
        frs.Range("B48").Interior.Color = vbCyan
        Me.CommandButton14.Enabled = True
        Me.Hide
        ws11.Visible = xlSheetVisible
        ws11.Activate
        
End Sub
Private Sub CommandButton15_Click()
    Dim ws6 As Worksheet
    Set ws6 = Sheets("Livestock Water Sheet")
    Dim ws7 As Worksheet
    Set ws7 = Sheets("Irrigation Water Sheet")
    Dim ws8 As Worksheet
    Set ws8 = Sheets("Domestic Water Sheet")
    Dim ws20 As Worksheet
    Set ws20 = Sheets("Storage Requirement Sheet")
    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
        
        ws20.Cells("2", "B").Value = ws7.Cells("32", "E").Value
        ws20.Cells("3", "B").Value = ws8.Cells("3", "B").Value
        ws20.Cells("4", "B").Value = ws6.Cells("16", "c").Value
        
    ws20.Calculate
    MsgBox "Total Water Demand is " & ws20.Cells("5", "B").Value & " Cubic Metres Per Day", vbInformation, "NOTE!"
    Me.CommandButton12.Enabled = True
    frs.Range("B36").Value = ws20.Cells("5", "B").Value
    frs.Range("B36").Interior.Color = vbCyan
    Me.Hide
    ws20.Visible = xlSheetVisible
    ws20.Activate
End Sub

Private Sub CommandButton30_Click()
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
    Dim frs As Worksheet
    Set frs = Sheets("Final Report Sheet")
        
        ws20.Cells("8", "B").Value = Me.TextBox13.Value
        ws20.Cells("7", "B").Value = Me.TextBox50.Value
        ws20.Cells("9", "B").Value = Me.TextBox30.Value
        
        frs.Range("B39").Value = ws20.Cells("8", "B").Value
        frs.Cells("39", "B").Interior.Color = vbCyan
        frs.Range("B40").Value = ws20.Cells("7", "B").Value
        frs.Cells("40", "B").Interior.Color = vbCyan
        frs.Range("B41").Value = ws20.Cells("9", "B").Value
        frs.Cells("41", "B").Interior.Color = vbCyan
        
    ws20.Calculate
    MsgBox "Total Storage Requirement for the Reservoir is " & ws20.Cells("10", "B").Value & " Cubic Metres", vbInformation, "NOTE!"
    
    
    If ws20.Cells("10", "B").Value > ws11.Cells("8", "B").Value Then
        MsgBox "The Reservoir can not Provide enough water given the Precipitation and runoff amounts. You Are Advised NOT TO PROCEED with the planning unless you adjust the water demand", vbCritical, "ERROR!"
        Me.CommandButton32.Enabled = True
        frs.Range("B50").Value = "FAIL"
        frs.Range("B50").Interior.Color = vbRed
        frs.Visible = xlSheetVisible
        frs.Activate
        Me.Hide
    Else
        MsgBox "The Reservoir Is Able To Provide enough water for the demand.", vbInformation, "NOTE!"
        Me.CommandButton32.Enabled = True
        frs.Range("B50").Value = "PASS"
        frs.Range("B50").Interior.Color = vbGreen
    End If
    
    frs.Range("B42").Value = ws20.Cells("10", "B").Value
    frs.Cells("42", "B").Interior.Color = vbCyan
    Me.Hide
    ws20.Visible = xlSheetVisible
    ws20.Activate
End Sub

Private Sub CommandButton29_Click()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = Sheets("Geotechnical Sheet")
    Set ws2 = Sheets("Geotechnical Sheet 2")
    
    If Me.MultiPage2.Value = 0 Then
        ws2.Visible = xlSheetVisible
        ws2.Activate
        Me.Hide
    ElseIf Me.MultiPage2.Value = 1 Then
        ws1.Visible = xlSheetVisible
        ws1.Activate
        Me.Hide
    End If
    
End Sub

Private Sub CommandButton33_Click()
    Dim calc As Worksheet
    Set calc = Sheets("CALCULATOR")
    Dim h As Long, z As Long, LastRow50 As Long, ws9 As Worksheet, frs As Worksheet
    Set ws9 = Sheets("HVA Table Sheet")
    LastRow50 = ws9.Range("F" & Rows.Count).End(xlUp).Row
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
    Set frs = Sheets("FInal Report Sheet")
        
        If ws20.Cells("10", "B").Value > ws11.Cells("8", "B").Value Then
            MsgBox "The Reservoir can not Provide enough water given the Precipitation and runoff amounts. You Are Advised NOT TO PROCEED with the planning unless you adjust the water demand.", vbCritical, "ERROR!"
        Else
            h = 2
            Do Until ws9.Cells(h, "F").Value >= ws20.Cells("10", "B").Value
                h = h + 1
            Loop
                ws9.Range("C" & h & ":F" & h).Interior.Color = vbGreen
                ws9.Range("AY2:AY" & h).Value = ws9.Cells(h, "D").Value
                ws9.Range("AZ2:AZ" & h).Value = ws9.Cells(h, "F").Value
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection(3).Name = "Corresponding Flooded Area"
                ActiveChart.FullSeriesCollection(3).Values = ws9.Range("AZ2:AZ" & h)
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection(4).Name = "Corresponding Volume"
                ActiveChart.FullSeriesCollection(4).Values = ws9.Range("AY2:AY" & h)
            MsgBox "The Optimum Water Depth for the reservoir is " & ws9.Cells(h, "C").Value & " Metres as highlighted in green.", vbInformation, "NOTE!"
            Me.CommandButton35.Enabled = True
            Me.TextBox54.Value = ws9.Cells(h, "C").Value
            calc.Range("B2").Value = ws9.Cells(h, "C").Value
            frs.Range("B53").Value = ws9.Cells(h, "C").Value
            frs.Cells("53", "B").Interior.Color = vbCyan
            frs.Range("B54").Value = ws9.Cells(h, "F").Value
            frs.Cells("54", "B").Interior.Color = vbCyan
        End If
        Me.Hide
        ws9.Visible = xlSheetVisible
        ws9.Activate
End Sub
