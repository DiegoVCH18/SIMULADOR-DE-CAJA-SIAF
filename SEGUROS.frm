VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SEGUROS 
   Caption         =   "SEGUROS - SIAF"
   ClientHeight    =   8625.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11235
   OleObjectBlob   =   "SEGUROS.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "SEGUROS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
TextBox12.Enabled = False
TextBox13.Enabled = False
TextBox14.Enabled = False
TextBox15.Enabled = False
TextBox16.Enabled = False
TextBox17.Enabled = False
TextBox18.Enabled = False
TextBox19.Enabled = False
TextBox20.Enabled = False
TextBox21.Enabled = False
TextBox22.Enabled = False
TextBox23.Enabled = False
TextBox24.Enabled = False
TextBox25.Enabled = False
TextBox26.Enabled = False
TextBox27.Enabled = False
TextBox28.Enabled = False
TextBox29.Enabled = False
TextBox30.Enabled = False
TextBox31.Enabled = False
TextBox32.Enabled = False
TextBox33.Enabled = False
TextBox34.Enabled = False
TextBox35.Enabled = False

Sheets("SEGURO VIDA").Select
Cells(29, 2) = CheckBox1.Caption
End Sub

Private Sub CheckBox2_Click()
TextBox12.Enabled = True
TextBox13.Enabled = True
TextBox14.Enabled = True
TextBox15.Enabled = True
TextBox16.Enabled = True
TextBox17.Enabled = True
TextBox18.Enabled = True
TextBox19.Enabled = True
TextBox20.Enabled = True
TextBox21.Enabled = True
TextBox22.Enabled = True
TextBox23.Enabled = True
TextBox24.Enabled = True
TextBox25.Enabled = True
TextBox26.Enabled = True
TextBox27.Enabled = True
TextBox28.Enabled = True
TextBox29.Enabled = True
TextBox30.Enabled = True
TextBox31.Enabled = True
TextBox32.Enabled = True
TextBox33.Enabled = True
TextBox34.Enabled = True
TextBox35.Enabled = True

Sheets("SEGURO VIDA").Select
Cells(29, 2) = CheckBox2.Caption
End Sub

Private Sub TextBox11_Change()
Sheets("SEGURO PT").Select
    Cells(87, 18) = TextBox11.Text
    Sheets("SEGURO VIDA").Select
    Cells(92, 18) = TextBox11.Text
End Sub

Private Sub TextBox12_Change()
Sheets("SEGURO VIDA").Select
                Cells(32, 2) = TextBox12.Text
End Sub



Private Sub TextBox12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox13_Change()
Sheets("SEGURO VIDA").Select
                Cells(32, 4) = TextBox13.Text
End Sub


Private Sub TextBox13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox14_Change()
Sheets("SEGURO VIDA").Select
                Cells(32, 6) = TextBox14.Text
End Sub

Private Sub TextBox14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox15_Change()
Sheets("SEGURO VIDA").Select
                Cells(32, 8) = TextBox15.Text
End Sub

Private Sub TextBox15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox16_Change()
Sheets("SEGURO VIDA").Select
                Cells(32, 10) = TextBox16.Text
End Sub

Private Sub TextBox16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox17_AfterUpdate()
TextBox17 = Format(TextBox17.Value, "#,###,###,##0.00%")
End Sub

Private Sub TextBox17_Change()
Sheets("SEGURO VIDA").Select
                Cells(32, 14) = TextBox17.Text
                TextBox37.Text = Format(Cells(36, 12), "#,###,###,##0.00%")
End Sub

Private Sub TextBox18_AfterUpdate()
TextBox18 = Format(TextBox18.Value, "#,###,###,##0.00%")
End Sub

Private Sub TextBox18_Change()
Sheets("SEGURO VIDA").Select
                Cells(33, 14) = TextBox18.Text
                TextBox37.Text = Format(Cells(36, 12), "#,###,###,##0.00%")
End Sub

Private Sub TextBox19_Change()
Sheets("SEGURO VIDA").Select
                Cells(33, 2) = TextBox19.Text
End Sub

Private Sub TextBox19_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox20_Change()
Sheets("SEGURO VIDA").Select
                Cells(33, 4) = TextBox20.Text
End Sub

Private Sub TextBox20_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox21_Change()
Sheets("SEGURO VIDA").Select
                Cells(33, 6) = TextBox21.Text
End Sub

Private Sub TextBox21_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox22_Change()
Sheets("SEGURO VIDA").Select
                Cells(33, 5) = TextBox22.Text
End Sub

Private Sub TextBox22_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox23_Change()
Sheets("SEGURO VIDA").Select
                Cells(33, 10) = TextBox23.Text
End Sub

Private Sub TextBox23_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox24_Change()
Sheets("SEGURO VIDA").Select
                Cells(34, 2) = TextBox24.Text
End Sub

Private Sub TextBox24_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox25_AfterUpdate()
TextBox25 = Format(TextBox25.Value, "#,###,###,##0.00%")
End Sub

Private Sub TextBox25_Change()
Sheets("SEGURO VIDA").Select
                Cells(34, 14) = TextBox25.Text
                TextBox37.Text = Format(Cells(36, 12), "#,###,###,##0.00%")
End Sub

Private Sub TextBox26_Change()
Sheets("SEGURO VIDA").Select
                Cells(34, 4) = TextBox26.Text
End Sub

Private Sub TextBox26_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox27_Change()
Sheets("SEGURO VIDA").Select
                Cells(34, 6) = TextBox27.Text
End Sub

Private Sub TextBox27_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox28_Change()
Sheets("SEGURO VIDA").Select
                Cells(34, 8) = TextBox28.Text
End Sub

Private Sub TextBox28_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox29_Change()
Sheets("SEGURO VIDA").Select
                Cells(34, 10) = TextBox29.Text
End Sub

Private Sub TextBox29_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox30_Change()
Sheets("SEGURO VIDA").Select
                Cells(35, 2) = TextBox30.Text
End Sub

Private Sub TextBox30_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox31_AfterUpdate()
TextBox31 = Format(TextBox31.Value, "#,###,###,##0.00%")
End Sub

Private Sub TextBox31_Change()
Sheets("SEGURO VIDA").Select
                Cells(35, 14) = TextBox31.Text
                TextBox37.Text = Format(Cells(36, 12), "#,###,###,##0.00%")
End Sub

Private Sub TextBox32_Change()
Sheets("SEGURO VIDA").Select
                Cells(35, 4) = TextBox32.Text
End Sub

Private Sub TextBox32_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox33_Change()
Sheets("SEGURO VIDA").Select
                Cells(35, 6) = TextBox33.Text
End Sub

Private Sub TextBox33_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox34_Change()
Sheets("SEGURO VIDA").Select
                Cells(35, 8) = TextBox34.Text
End Sub

Private Sub TextBox34_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox35_Change()
Sheets("SEGURO VIDA").Select
                Cells(35, 10) = TextBox35.Text
End Sub

Private Sub TextBox35_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub TextBox7_Change()
Sheets("SEGURO PT").Select
                Cells(87, 16) = TextBox7.Text

End Sub

Private Sub TextBox9_Change()
Sheets("SEGURO PT").Select
    Cells(87, 19) = TextBox9.Text
    Sheets("SEGURO VIDA").Select
    Cells(92, 19) = TextBox9.Text

End Sub

'propiedad intelectual ing. Diego Armando Vásquez Chávez DNI: 44113825
Private Sub ComboBox1_Change()
TextBox1.Text = ""
If ComboBox1.Text = "RUC" Then
    TextBox1.MaxLength = 11
    Else
    CLIENTE.Show
    TextBox1.MaxLength = 8
End If
End Sub

Private Sub ComboBox2_Change()

If ComboBox2.Text = "Seguro contra robos de Tarjetas Plus" Then
CommandButton1.Visible = True
CommandButton3.Visible = False
    If Frame2.Visible = True Then
    Frame2.Visible = False
    Sheets("SEGURO PT").Select
    TextBox8.Text = Cells(87, 4)
    End If
Else
CommandButton1.Visible = False
CommandButton3.Visible = True
Frame2.Visible = True
Sheets("SEGURO VIDA").Select
TextBox8.Text = Cells(92, 4)
End If

End Sub

Private Sub ComboBox4_Change()
Sheets("SEGURO VIDA").Select
    Cells(90, 4) = ComboBox4.Text
Sheets("SEGURO PT").Select
    Cells(85, 4) = ComboBox4.Text
    
End Sub

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

                             
                Dim NombreArchivo, RutaArchivo As String
                    Sheets("SEGURO PT").Select
                    FECHA = Label1.Caption
                    NombreArchivo = "SOLICITUD SEGURO PT" & " " & CStr(Format(Date, "dd-mm")) & " " & CStr(Format(Time, "hh-mm-ss"))
                    RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
                    ActiveSheet.Range("A2:N110").Select
                    Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaArchivo, _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, OpenAfterPublish:=True
                    
                    
                    'Dim Ruta As String, nombre As String
'Ruta = ThisWorkbook.Path
'nombre = Ruta & "\" & ActiveSheet.Name 'Nombre de hoja
    'Range("C1:E29").Select
    'Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nombre & ".pdf", _
'Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'OpenAfterPublish:=True 'Si quieres que el archivo se abra luego de creado, cambia False por True al final
'Range("A1").Select 'celda final selecionada
                    
                     With ActiveSheet.PageSetup
                            .LeftHeader = ""
                            .CenterHeader = ""
                            .RightHeader = ""
                            .LeftFooter = ""
                            .CenterFooter = ""
                            .RightFooter = ""
                            .LeftMargin = Application.InchesToPoints(0.1)
                            .RightMargin = Application.InchesToPoints(0.1)
                            .TopMargin = Application.InchesToPoints(0)
                            .BottomMargin = Application.InchesToPoints(0.1)
                            .HeaderMargin = Application.InchesToPoints(0.1)
                            .FooterMargin = Application.InchesToPoints(0.1)
                            .PrintHeadings = False
                            .PrintGridlines = False
                            .PrintComments = xlPrintNoComments
                            .CenterHorizontally = True
                            .CenterVertically = False
                            .Orientation = xlPortrait
                            .Draft = False
                            .FirstPageNumber = xlAutomatic
                            .Order = xlDownThenOver
                            .BlackAndWhite = False
                            .Zoom = False
                            .FitToPagesWide = 1
                          
                            .PrintErrors = xlPrintErrorsDisplayed
                            .OddAndEvenPagesHeaderFooter = False
                            .DifferentFirstPageHeaderFooter = False
                            .ScaleWithDocHeaderFooter = True
                            .AlignMarginsHeaderFooter = False
                            .EvenPage.LeftHeader.Text = ""
                            .EvenPage.CenterHeader.Text = ""
                            .EvenPage.RightHeader.Text = ""
                            .EvenPage.LeftFooter.Text = ""
                            .EvenPage.CenterFooter.Text = ""
                            .EvenPage.RightFooter.Text = ""
                            .FirstPage.LeftHeader.Text = ""
                            .FirstPage.CenterHeader.Text = ""
                            .FirstPage.RightHeader.Text = ""
                            .FirstPage.LeftFooter.Text = ""
                            .FirstPage.CenterFooter.Text = ""
                            .FirstPage.RightFooter.Text = ""
                        End With
                       Application.ScreenUpdating = False
                        Application.Visible = True
  
                                                 
                
        
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Private Sub CommandButton2_Click()
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = False
Sheets("ULTIMO REGISTRO").Visible = False
Sheets("TIPO DE CAMBIO").Visible = False
Sheets("ULTIMA CUENTA").Visible = False
Sheets("BASE CUENTAS").Visible = False

ComboBox1.Text = ""
             ComboBox2.Text = ""
             ComboBox4.Text = ""
             TextBox1.Text = ""
             TextBox2.Text = ""

           
     
     
             TextBox8.Text = ""
         
            
         
             
             
             Sheets("SOLICITUD TC").Select
             Range( _
        "D19:E19,G18,I19:J19,B23:E23,F23:I23,J23:M23,D25:E25,G25,J25:M25,B29:M29,B33:E33,F33:I33,J33:M33,B37,C37:E37,F37:I37,J37:M37,D39:G39,K39:M39,B43:E43,F43:I43,J43:M43" _
        ).Select
    Range("J43").Activate
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
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
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    Union(Range( _
        "E63:G63,H63,I63:K63,L63:M63,D67,F67:G67,L67:M67,B71:E71,F71:I71,J71:M71,D73:E73,H73:I73,K73:M73,B77:D77,E77:G77,H77:J77,D19:E19,G18,I19:J19,B23:E23,F23:I23,J23:M23,D25:E25,G25,J25:M25,B29:M29,B33:E33,F33:I33,J33:M33,B37,C37:E37,F37:I37" _
        ), Range( _
        "J37:M37,D39:G39,K39:M39,B43:E43,F43:I43,J43:M43,B49:C49,D49:F49,G49:M49,B53:M53,B57:E57,F57:I57,J57:M57,D59:E59,H59:M59,B63:D63" _
        )).Select
    Range("H77").Activate
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 87
    Union(Range( _
        "E63:G63,H63,I63:K63,L63:M63,D67,F67:G67,L67:M67,B71:E71,F71:I71,J71:M71,D73:E73,H73:I73,K73:M73,B77:D77,E77:G77,H77:J77,D90:E90,L90:M90,B94:D94,E94:M94,D118:J118,D120:E120,G120,D122:J122,D124:E124,G124,D19:E19,G18,I19:J19,B23:E23,F23:I23,J23:M23" _
        ), Range( _
        "D25:E25,G25,J25:M25,B29:M29,B33:E33,F33:I33,J33:M33,B37,C37:E37,F37:I37,J37:M37,D39:G39,K39:M39,B43:E43,F43:I43,J43:M43,B49:C49,D49:F49,G49:M49,B53:M53,B57:E57,F57:I57,J57:M57,D59:E59,H59:M59,B63:D63" _
        )).Select
    Range("G124").Activate
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-96
    Range("D19:E19").Select
    
Me.Hide

End Sub

Private Sub CommandButton3_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

                             
                Dim NombreArchivo, RutaArchivo As String
                    Sheets("SEGURO VIDA").Select
                    FECHA = Label1.Caption
                    NombreArchivo = "SOLICITUD SEGURO MÚLTIPLE" & " " & CStr(Format(Date, "dd-mm")) & " " & CStr(Format(Time, "hh-mm-ss"))
                    RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
                    ActiveSheet.Range("A2:N110").Select
                    Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaArchivo, _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, OpenAfterPublish:=True
                    
                    
                    'Dim Ruta As String, nombre As String
'Ruta = ThisWorkbook.Path
'nombre = Ruta & "\" & ActiveSheet.Name 'Nombre de hoja
    'Range("C1:E29").Select
    'Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nombre & ".pdf", _
'Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'OpenAfterPublish:=True 'Si quieres que el archivo se abra luego de creado, cambia False por True al final
'Range("A1").Select 'celda final selecionada
                    
                     With ActiveSheet.PageSetup
                            .LeftHeader = ""
                            .CenterHeader = ""
                            .RightHeader = ""
                            .LeftFooter = ""
                            .CenterFooter = ""
                            .RightFooter = ""
                            .LeftMargin = Application.InchesToPoints(0.1)
                            .RightMargin = Application.InchesToPoints(0.1)
                            .TopMargin = Application.InchesToPoints(0)
                            .BottomMargin = Application.InchesToPoints(0.1)
                            .HeaderMargin = Application.InchesToPoints(0.1)
                            .FooterMargin = Application.InchesToPoints(0.1)
                            .PrintHeadings = False
                            .PrintGridlines = False
                            .PrintComments = xlPrintNoComments
                            .CenterHorizontally = True
                            .CenterVertically = False
                            .Orientation = xlPortrait
                            .Draft = False
                            .FirstPageNumber = xlAutomatic
                            .Order = xlDownThenOver
                            .BlackAndWhite = False
                            .Zoom = False
                            .FitToPagesWide = 1
                          
                            .PrintErrors = xlPrintErrorsDisplayed
                            .OddAndEvenPagesHeaderFooter = False
                            .DifferentFirstPageHeaderFooter = False
                            .ScaleWithDocHeaderFooter = True
                            .AlignMarginsHeaderFooter = False
                            .EvenPage.LeftHeader.Text = ""
                            .EvenPage.CenterHeader.Text = ""
                            .EvenPage.RightHeader.Text = ""
                            .EvenPage.LeftFooter.Text = ""
                            .EvenPage.CenterFooter.Text = ""
                            .EvenPage.RightFooter.Text = ""
                            .FirstPage.LeftHeader.Text = ""
                            .FirstPage.CenterHeader.Text = ""
                            .FirstPage.RightHeader.Text = ""
                            .FirstPage.LeftFooter.Text = ""
                            .FirstPage.CenterFooter.Text = ""
                            .FirstPage.RightFooter.Text = ""
                        End With
                       Application.ScreenUpdating = False
                        Application.Visible = True
                          Unload Me
                       
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox2.Text = UCase(TextBox2.Text)
End Sub

Private Sub UserForm_activate()
Application.ScreenUpdating = False
Sheets("SOLICITUD TC").Cells(9, 11) = MENU.Label8.Caption
Sheets("SOLICITUD TC").Cells(130, 5) = MENU.TextBox1.Text
Sheets("SOLICITUD TC").Cells(130, 8) = MENU.TextBox2.Text
Sheets("SOLICITUD TC").Cells(130, 12) = MENU.TextBox4.Text
ComboBox1.AddItem ("DNI")
ComboBox1.AddItem ("CE")
ComboBox4.AddItem ("Plan Mensual")
ComboBox4.AddItem ("Plan Anual")
ComboBox2.AddItem ("Seguro contra robos de Tarjetas Plus")
ComboBox2.AddItem ("Seguro de Vida y Accidentes")


Sheets("SEGURO PT").Select
    Cells(87, 16) = TextBox8.Text
Sheets("SEGURO VIDA").Select
    Cells(92, 16) = TextBox8.Text

Sheets("SEGURO PT").Select
    
Sheets("SEGURO VIDA").Select
  
    Frame1.Visible = True
    Frame2.Visible = True
 Sheets("SEGURO VIDA").Select
                Cells(32, 2) = ""
                Cells(32, 4) = ""
                Cells(32, 6) = ""
                Cells(32, 8) = ""
                Cells(32, 10) = ""
                Cells(32, 14) = ""
                Cells(33, 2) = ""
                Cells(33, 4) = ""
                Cells(33, 6) = ""
                Cells(33, 8) = ""
                Cells(33, 10) = ""
                Cells(33, 14) = ""
                Cells(34, 2) = ""
                Cells(34, 4) = ""
                Cells(34, 6) = ""
                Cells(34, 8) = ""
                Cells(34, 10) = ""
                Cells(34, 14) = ""
                Cells(35, 2) = ""
                Cells(35, 4) = ""
                Cells(35, 6) = ""
                Cells(35, 8) = ""
                Cells(35, 10) = ""
                Cells(35, 14) = ""
TextBox8.Visible = True


CommandButton1.Visible = False
CommandButton3.Visible = False

Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Application.Visible = False
Application.ScreenUpdating = False
End Sub
Private Sub TextBox1_Keypress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub
Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub


