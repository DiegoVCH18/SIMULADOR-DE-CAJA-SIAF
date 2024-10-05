VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RECLAMOS 
   Caption         =   "HOJA DE RECLAMACIÓN - SIAF"
   ClientHeight    =   9855.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   OleObjectBlob   =   "RECLAMOS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RECLAMOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

If CheckBox1.Value = False Then
TextBox12.Visible = False
Else
TextBox12.Visible = True
End If
PRODUCTOS.Show
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
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(16, 3) = ComboBox2.Text
End Sub


Private Sub ComboBox3_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(49, 2) = ComboBox3.Text
End Sub

Private Sub ComboBox4_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(62, 8) = ComboBox4.Text
End Sub

Private Sub ComboBox5_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(56, 2) = ComboBox5.Text
End Sub

Private Sub ComboBox6_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(98, 2) = ComboBox6.Text
End Sub

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Sheets("HOJA DE RECLAMO").Select
    If TextBox11.Text = "" Then
    MsgBox "Completar Hoja de Reclamación", , "SIAF v 1.2.0"
        Else
        If ComboBox2.Text = "" Or ComboBox5.Text = "" Or ComboBox6.Text = "" Then
        MsgBox "Completar Hoja de Reclamación", , "SIAF v 1.2.0"
            Else
     
                             
                Dim NombreArchivo, RutaArchivo As String
                    Sheets("HOJA DE RECLAMO").Select
                    FECHA = Label1.Caption
                    NombreArchivo = "HOJA DE RECLAMACIÓN" & " " & CStr(Format(Date, "dd-mm")) & " " & CStr(Format(Time, "hh-mm-ss"))
                    RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
                    ActiveSheet.Range("A2:N150").Select
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
                            
                                      
            End If
                End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
 Sheets("HOJA DE RECLAMO").Select
    ActiveWindow.SmallScroll Down:=6
    Range("B49:M49").Select
    ActiveWindow.SmallScroll Down:=6
    Range("B49:M49,B52:F52,B56:M56,K53,I52:M52,D62:E62,H62:I62,K62:M62,B67:M77"). _
        Select
    Range("B67").Activate
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
    Range( _
        "B49:M49,B52:F52,B56:M56,K53,I52:M52,D62:E62,H62:I62,K62:M62,B67:M77,B82:M91,B98:D98" _
        ).Select
    Range("B98").Activate
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
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-84
    Range("C13:F13").Select

End Sub

Private Sub TextBox10_Change()
Application.ScreenUpdating = False
TextBox10.MaxLength = 10
largo_entrada = Len(Me.TextBox10)
Select Case largo_entrada
Case 2
Me.TextBox10.Value = Me.TextBox10.Value & "/"
Case 5
Me.TextBox10.Value = Me.TextBox10.Value & "/"
End Select

Sheets("HOJA DE RECLAMO").Cells(62, 4) = TextBox10.Text
End Sub


Private Sub TextBox8__KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox11_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(67, 2) = TextBox11.Text
End Sub

Private Sub TextBox11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox11.Text = UCase(TextBox11.Text)
End Sub

Private Sub TextBox12_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(82, 2) = TextBox12.Text
End Sub

Private Sub TextBox12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox12.Text = UCase(TextBox12.Text)
End Sub

Private Sub TextBox7_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(52, 2) = TextBox7.Text
End Sub

Private Sub TextBox8_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(62, 11) = TextBox8.Text
End Sub

Private Sub TextBox9_Change()
Application.ScreenUpdating = False
Sheets("HOJA DE RECLAMO").Cells(52, 9) = TextBox9.Text
End Sub

Private Sub UserForm_activate()
Application.ScreenUpdating = False
Sheets("SOLICITUD TC").Cells(9, 11) = MENU.Label8.Caption
Sheets("SOLICITUD TC").Cells(130, 5) = MENU.TextBox1.Text
Sheets("SOLICITUD TC").Cells(130, 8) = MENU.TextBox2.Text
Sheets("SOLICITUD TC").Cells(130, 12) = MENU.TextBox4.Text

ComboBox1.AddItem ("DNI")
ComboBox1.AddItem ("CE")
ComboBox2.AddItem ("QUEJA")
ComboBox2.AddItem ("RECLAMO")

ComboBox3.AddItem ("CUENTA DE AHORRO")
ComboBox3.AddItem ("CUENTA A PLAZO FIJO")
ComboBox3.AddItem ("TARJETAS DE CRÉDITO")
ComboBox3.AddItem ("BANCA - SEGUROS (SEGUROS VENDIDOS EN LOS CANALES DEL SISTEMA FINANCIERO)")
ComboBox3.AddItem ("CUENTA CORRIENTE")
ComboBox3.AddItem ("CRÉDITO PERSONAL")
ComboBox3.AddItem ("ATENCIÓN AL PÚBLICO (NO RELACIONADO A LAS OPERACIONES O PRODUCTOS OFRECIDOS POR LA EMPRESA)")
ComboBox3.AddItem ("CRÉDITO HIPOTECARIO PARA VIVIENDA")
ComboBox3.AddItem ("CUENTA CTS")
ComboBox3.AddItem ("TARJETA DE DÉBITO")
ComboBox3.AddItem ("GIROS")

ComboBox4.AddItem ("MN S/")
ComboBox4.AddItem ("ME $")

ComboBox5.AddItem ("COBROS INDEBIDOS DE INTERESES, COMISIONES, GASTOS Y TRIBUTOS (TALES COMO SEGUROS, ITF, ENTRE OTROS CARGOS, SEGÚN CORRESPONDA)")
ComboBox5.AddItem ("DEMORAS O INCUMPLIMIENTOS DE ENVÍO DE CORRESPONDENCIA (ESTADOS DE CUENTA, OTROS)")
ComboBox5.AddItem ("DISCONFORMIDAD POR NOTIFICACIONES DIRIGIDAS A TERCERAS PERSONAS")
ComboBox5.AddItem ("ERROR EN LOS DATOS DEL USUARIO REGISTRADO EN LA EMPRESA")
ComboBox5.AddItem ("FALLAS DEL SISTEMA INFORMÁTICO QUE DIFICULTAN OPERACIONES Y SERVICIOS.")
ComboBox5.AddItem ("INADECUADA ATENCIÓN AL USUARIO - PROBLEMAS EN LA CALIDAD DEL SERVICIO")
ComboBox5.AddItem ("OPERACIONES NO RECONOCIDAS (CONSUMOS, DISPOSICIONES, RETIROS, CARGOS, ABONOS Y SOBREGIROS, SEGÚN CORRESPONDA)")
ComboBox5.AddItem ("PROBLEMAS PRESENTADOS CON LA TARJETA DE CRÉDITO O DÉBITO (RETENIDA, NO EMITIDA, NO ENTREGADA A TIEMPO, DESACTIVADA, BLOQUEADA, ANULADA, SUSPENDIDA, CANCELADA)")
ComboBox5.AddItem ("TRANSACCIONES NO PROCESADAS / MAL REALIZADAS")
ComboBox5.AddItem ("OTROS")

ComboBox6.AddItem ("DIRECCION DE DOMICILIO")
ComboBox6.AddItem ("CORREO ELECTRÓNICO")
ComboBox6.AddItem ("OFICINA EMISORA")
ComboBox6.AddItem ("FUNCIONARIO DE NEGOCIOS")


Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Application.Visible = True
Application.ScreenUpdating = False
End Sub
