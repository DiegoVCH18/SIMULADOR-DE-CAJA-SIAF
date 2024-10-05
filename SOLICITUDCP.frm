VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SOLICITUDCP 
   Caption         =   "SOLICITUD DE CRÉDITO PERSONAL - SIAF"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14100
   OleObjectBlob   =   "SOLICITUDCP.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "SOLICITUDCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub ComboBox10_Change()
Sheets("SOLICITUD TC").Cells(124, 4) = ComboBox10.Text
End Sub

Private Sub ComboBox11_Change()
Sheets("SOLICITUD CP").Cells(84, 2) = ComboBox11.Text
End Sub

Private Sub ComboBox2_Change()
Sheets("SOLICITUD CP").Select
Cells(84, 5) = ComboBox2.Text

End Sub

Private Sub ComboBox4_Change()
If ComboBox4.Text = "MN S/" Then
    TextBox5.Text = 0
Else
    TextBox5.Text = 1
End If
End Sub

Private Sub ComboBox5_Change()
Sheets("SOLICITUD CP").Cells(80, 3) = ComboBox5.Text
End Sub

Private Sub ComboBox6_Change()
Sheets("SOLICITUD CP").Cells(80, 8) = ComboBox6.Text
End Sub

Private Sub ComboBox8_Change()
Sheets("SOLICITUD TC").Cells(94, 5) = ComboBox8.Text
End Sub

Private Sub ComboBox9_Change()
Sheets("SOLICITUD TC").Cells(120, 4) = ComboBox9.Text
End Sub

Private Sub CommandButton1_Click()
If ComboBox1 = "" Or TextBox1 = "" Or TextBox2 = "" Or ComboBox2 = "" Or ComboBox4 = "" Then
    MsgBox "Completar datos "
    Else
        Sheets("ULTIMA CUENTA").Select
                Cells(2, 1) = TextBox8.Value + "-" + TextBox10.Value + "-" + TextBox11.Value + "-" + TextBox9.Value
                Cells(2, 2) = ComboBox1.Text
                Cells(2, 3) = TextBox1.Text
                Cells(2, 4) = TextBox2.Text
                Cells(2, 5) = ComboBox2.Text
                Cells(2, 6) = ComboBox4.Text
                Cells(2, 7) = TextBox7.Value + "-" + TextBox4.Value + "-" + TextBox5.Value + "-" + TextBox6.Value
                                   
               'Codigo obtenido del grabador de macros
                Sheets("BASE CUENTAS").Select
                Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                With Selection.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Sheets("ULTIMA CUENTA").Select
                Range("A2:O2").Select
                Selection.Copy
                Range("A2").Select
                Sheets("BASE CUENTAS").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                Sheets("BASE CUENTAS").Select
                MsgBox " Registrado Correctamente"
                ComboBox1.Text = ""
             ComboBox2.Text = ""
             ComboBox4.Text = ""
             TextBox1.Text = ""
             TextBox2.Text = ""
             TextBox7.Text = ""
             TextBox4.Text = ""
             TextBox5.Text = ""
             TextBox6.Text = ""
             TextBox8.Text = ""
             TextBox10.Text = ""
             TextBox11.Text = ""
             TextBox9.Text = ""
             Me.Hide
                End If
                           

End Sub

Private Sub CommandButton2_Click()
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = False
Sheets("ULTIMO REGISTRO").Visible = False
Sheets("TIPO DE CAMBIO").Visible = False
Sheets("ULTIMA CUENTA").Visible = False
Sheets("BASE CUENTAS").Visible = False
Sheets("BUSC TARJETA").Visible = False
ComboBox1.Text = ""
             ComboBox2.Text = ""
             ComboBox4.Text = ""
             TextBox1.Text = ""
             TextBox2.Text = ""
             TextBox7.Text = ""
             TextBox4.Text = ""
             TextBox5.Text = ""
             TextBox6.Text = ""
             TextBox8.Text = ""
             TextBox10.Text = ""
             TextBox11.Text = ""
             TextBox9.Text = ""
Me.Hide

End Sub

Private Sub CommandButton3_Click()
If ComboBox1 = "" Or TextBox1 = "" Or TextBox2 = "" Or ComboBox2 = "" Or ComboBox4 = "" Then
    MsgBox "Completar datos "
    Else
    Frame1.Visible = True
    TextBox4.Visible = True
TextBox7.Visible = True
TextBox4.Visible = True
TextBox5.Visible = True
TextBox6.Visible = True
TextBox8.Visible = True
TextBox10.Visible = True
TextBox11.Visible = True
TextBox9.Visible = True
TextBox11.MaxLength = 4
TextBox9.MaxLength = 4
If TextBox4.Text <> "" Then
MsgBox "Cuenta creada"
Else
If ComboBox2.Text = "2.CUENTA CORRIENTE" Then
    TextBox4.Text = Int((9999999 * Rnd) + 1000000)
    Else
    TextBox4.Text = Int((99999999 * Rnd) + 10000000)
End If
TextBox6.Text = Int((99 * Rnd) + 10)
End If
If ComboBox2.Text = "4.CTA. PLAZO" Then
Label7.Visible = False
TextBox8.Visible = False
TextBox10.Visible = False
TextBox11.Visible = False
TextBox9.Visible = False
End If

End If

End Sub

Private Sub CommandButton4_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Definir las hojas de trabajo
    Set wsCartillaCuenta = ThisWorkbook.Sheets("SOLICITUD CP")
    Set wsTipoCambio = ThisWorkbook.Sheets("TIPO DE CAMBIO")
    
    ' Encontrar la próxima fila vacía en la hoja Tipo de Cambio
    nextRow = wsTipoCambio.Cells(wsTipoCambio.Rows.Count, "B").End(xlUp).Row + 1
    
    ' Copiar los datos desde la hoja Cartilla Cuenta a la hoja Tipo de Cambio
    wsCartillaCuenta.Range("S10:S14").Copy
    
    ' Pegar los datos en la próxima fila vacía de la hoja Tipo de Cambio en la columna B
    wsTipoCambio.Cells(nextRow, "B").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    
    ' Obtener la última columna utilizada en la hoja Tipo de Cambio
    lastColumn = wsTipoCambio.Cells(nextRow, wsTipoCambio.Columns.Count).End(xlToLeft).Column
    
    ' Obtener la última fila utilizada en la columna A de la Tabla2
    lastRowTabla2 = wsTipoCambio.Cells(wsTipoCambio.Rows.Count, "A").End(xlUp).Row
    
    ' Continuar con la numeración en la columna A de la hoja Tipo de Cambio
    wsTipoCambio.Cells(nextRow, "A").Value = wsTipoCambio.Cells(lastRowTabla2, "A").Value + 1



    If TextBox15.Text = "" Then
    MsgBox "Completar Solicitud de Crédito Personal", , "SIAF v 1.2.0"
        Else
        If ComboBox2.Text = "" Or ComboBox5.Text = "" Or ComboBox11.Text = "" Then
        MsgBox "Completar Solicitud de Crédito Personal", , "SIAF v 1.2.0"
            Else
     
                             
                Dim NombreArchivo, RutaArchivo As String
                    Sheets("SOLICITUD CP").Select
                    FECHA = Label1.Caption
                    NombreArchivo = "SOLICITUD DE CRÉDITO PERSONAL" & " " & CStr(Format(Date, "dd-mm")) & " " & CStr(Format(Time, "hh-mm-ss"))
                    RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
                    ActiveSheet.Range("A2:N304").Select
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
                    Application.Visible = False
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
                            
                                      
            End If
                End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

  
End Sub

Private Sub CommandButton5_Click()
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

Private Sub CommandButton6_Click()
PROCESO.Show

End Sub

Private Sub Frame2_Click()
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label28_Click()

End Sub

Private Sub TextBox12_Change()

End Sub

Private Sub TextBox13_Change()
Sheets("SOLICITUD TC").Cells(110, 2) = TextBox13.Text
End Sub

Private Sub TextBox14_Change()
Sheets("SOLICITUD TC").Cells(110, 5) = TextBox14.Text
End Sub

Private Sub TextBox15_Change()
Sheets("SOLICITUD TC").Cells(114, 7) = TextBox15.Text
Label24 = TextBox15.Text
End Sub

Private Sub TextBox15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox15.Text = UCase(TextBox15.Text)
End Sub

Private Sub TextBox16_Change()
Sheets("SOLICITUD TC").Cells(118, 4) = TextBox16.Text
End Sub

Private Sub TextBox16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox16.Text = UCase(TextBox16.Text)
End Sub

Private Sub TextBox17_Change()

End Sub

Private Sub TextBox17_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox18_Change()
Sheets("SOLICITUD TC").Cells(122, 4) = TextBox18.Text
End Sub

Private Sub TextBox18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox18.Text = UCase(TextBox18.Text)
End Sub

Private Sub TextBox19_Change()

End Sub

Private Sub TextBox19_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox2.Text = UCase(TextBox2.Text)
End Sub

Private Sub TextBox22_Change()
Sheets("SOLICITUD CP").Cells(139, 9) = TextBox22.Text
End Sub

Private Sub UserForm_activate()
Application.ScreenUpdating = False
Sheets("SOLICITUD TC").Cells(9, 11) = MENU.Label8.Caption
Sheets("SOLICITUD TC").Cells(130, 5) = MENU.TextBox1.Text
Sheets("SOLICITUD TC").Cells(130, 8) = MENU.TextBox2.Text
Sheets("SOLICITUD TC").Cells(130, 12) = MENU.TextBox4.Text

ComboBox1.AddItem ("DNI")
ComboBox1.AddItem ("CE")
ComboBox2.AddItem ("Dólares $")
ComboBox2.AddItem ("Soles S/")
ComboBox5.AddItem ("1")
ComboBox5.AddItem ("2")
ComboBox5.AddItem ("3")
ComboBox5.AddItem ("4")
ComboBox5.AddItem ("5")
ComboBox5.AddItem ("6")
ComboBox5.AddItem ("7")
ComboBox5.AddItem ("8")
ComboBox5.AddItem ("9")
ComboBox5.AddItem ("10")
ComboBox5.AddItem ("11")
ComboBox5.AddItem ("12")
ComboBox5.AddItem ("13")
ComboBox5.AddItem ("14")
ComboBox5.AddItem ("15")
ComboBox5.AddItem ("16")
ComboBox5.AddItem ("17")
ComboBox5.AddItem ("18")
ComboBox5.AddItem ("19")
ComboBox5.AddItem ("20")
ComboBox5.AddItem ("21")
ComboBox5.AddItem ("22")
ComboBox5.AddItem ("23")
ComboBox5.AddItem ("24")
ComboBox5.AddItem ("25")
ComboBox5.AddItem ("26")
ComboBox5.AddItem ("27")
ComboBox5.AddItem ("28")
ComboBox5.AddItem ("29")
ComboBox5.AddItem ("30")
ComboBox5.AddItem ("31")
ComboBox6.AddItem ("6")
ComboBox6.AddItem ("12")
ComboBox6.AddItem ("18")
ComboBox6.AddItem ("24")
ComboBox6.AddItem ("36")
ComboBox6.AddItem ("48")
ComboBox6.AddItem ("60")


ComboBox11.AddItem ("Físico")
ComboBox11.AddItem ("Virtual")

Frame3.Visible = False

Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Application.Visible = True
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





