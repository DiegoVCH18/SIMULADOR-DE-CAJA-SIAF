VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LAVA 
   Caption         =   "Registro de operaciones en efectivo de mayor cuantía  - SIAF"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   OleObjectBlob   =   "LAVA.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "LAVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botonciiu2_Click()
CIIU2.Show
End Sub

Private Sub botonciiu3_Click()
CIIU3.Show
End Sub

Private Sub CheckBox2_Click()
ComboBox4.Text = ComboBox3.Text
TextBox34.Text = TextBox24.Text
TextBox38.Text = TextBox28.Text
TextBox33.Text = TextBox23.Text
TextBox35.Text = TextBox25.Text
TextBox36.Text = TextBox26.Text
TextBox37.Text = TextBox27.Text
TextBox42.Text = TextBox32.Text
TextBox39.Text = TextBox29.Text
TextBox40.Text = TextBox30.Text
TextBox41.Text = TextBox31.Text
Sheets("LAVA").Select
Cells(26, 3) = ComboBox4.Text
Cells(26, 5) = TextBox34.Text
Cells(27, 3) = TextBox38.Text
Cells(28, 3) = TextBox33.Text
Cells(29, 3) = TextBox35.Text
Cells(29, 6) = TextBox36.Text
Cells(29, 9) = TextBox37.Text
Cells(30, 3) = TextBox42.Text
Cells(30, 6) = TextBox39.Text
Cells(30, 9) = TextBox40.Text
Cells(31, 3) = TextBox41.Text

End Sub

Private Sub CheckBox3_Click()
ComboBox5.Text = ComboBox3.Text
TextBox44.Text = TextBox24.Text
TextBox48.Text = TextBox28.Text
TextBox43.Text = TextBox23.Text
TextBox45.Text = TextBox25.Text
TextBox46.Text = TextBox26.Text
TextBox47.Text = TextBox27.Text
TextBox52.Text = TextBox32.Text
TextBox49.Text = TextBox29.Text
TextBox50.Text = TextBox30.Text
TextBox51.Text = TextBox31.Text
Sheets("LAVA").Select
Cells(35, 3) = ComboBox5.Text
Cells(35, 5) = TextBox44.Text
Cells(36, 3) = TextBox48.Text
Cells(37, 3) = TextBox43.Text
Cells(38, 3) = TextBox45.Text
Cells(38, 6) = TextBox46.Text
Cells(38, 9) = TextBox47.Text
Cells(39, 3) = TextBox52.Text
Cells(39, 6) = TextBox49.Text
Cells(39, 9) = TextBox50.Text
Cells(40, 3) = TextBox51.Text
End Sub

Private Sub ComboBox3_Change()
Sheets("LAVA").Select
Cells(16, 3) = ComboBox3.Text

End Sub

Private Sub ComboBox4_Change()
Sheets("LAVA").Select
Cells(26, 3) = ComboBox4.Text
If ComboBox4.Text = "RUC" Then
Label84.Caption = "Fecha constitución"
Label86.Caption = "Giro"
botonciiu2.Visible = True
CommandButton18.Visible = False
End If
End Sub

Private Sub ComboBox5_Change()
Sheets("LAVA").Select
Cells(35, 3) = ComboBox5.Text
If ComboBox5.Text = "RUC" Then
Label95.Caption = "Fecha constitución"
Label97.Caption = "Giro"
botonciiu3.Visible = True
CommandButton20.Visible = False
End If
End Sub

Private Sub CommandButton13_Click()
Sheets("LAVA").Select
NombreArchivo = "FORMULARIO LAVA" & " " & CStr(Format(Date, "dd-mm")) & " " & CStr(Format(Time, "hh-mm-ss"))
RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
                    ActiveSheet.Range("A1:L64").Select
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
                        DEPO.TextBox15 = "COMPLETO"
                        CANC.TextBox15 = "COMPLETO"
                        RETI.TextBox15 = "COMPLETO"
                        CHPA.TextBox15 = "COMPLETO"
                        COBR.TextBox15 = "COMPLETO"
                        EMIS.TextBox15 = "COMPLETO"
                        PAGO.TextBox15 = "COMPLETO"
                        Unload Me
                        ActiveWindow.SmallScroll Down:=-30
    Range("C16").Select
    ActiveWindow.SmallScroll Down:=6
    Union(Range( _
        "C40:J41,C16,E16,C17:J17,C18:J18,C19:D19,F19,I19,C20,C21,F20,I20,C26,E26:F26,C27:J27,C28:J28,C29:D29,F29,I29,C30,F30,I30,C31:J32,C35,E35:F35,C36:J36,C37:J37,C38:D38,C39,F38,F39,I38" _
        ), Range("I39")).Select
    Range("C40").Activate
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
    Union(Range( _
        "C40:J41,B43:J45,B47:J49,C16,E16,C17:J17,C18:J18,C19:D19,F19,I19,C20,C21,F20,I20,C26,E26:F26,C27:J27,C28:J28,C29:D29,F29,I29,C30,F30,I30,C31:J32,C35,E35:F35,C36:J36,C37:J37,C38:D38,C39,F38" _
        ), Range("F39,I38,I39")).Select
    Range("B47").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-27
    Range("C16").Select
End Sub

Private Sub CommandButton14_Click()

ComboBox3.Text = ""
ComboBox4.Text = ""
ComboBox5.Text = ""
TextBox34.Text = ""
TextBox24.Text = ""
TextBox38.Text = ""
 TextBox28.Text = ""
TextBox33.Text = ""
TextBox23.Text = ""
TextBox35.Text = ""
TextBox25.Text = ""
TextBox36.Text = ""
TextBox26.Text = ""
TextBox37.Text = ""
TextBox27.Text = ""
TextBox42.Text = ""
TextBox32.Text = ""
TextBox39.Text = ""
TextBox29.Text = ""
TextBox40.Text = ""
TextBox30.Text = ""
TextBox41.Text = ""
TextBox31.Text = ""
TextBox44.Text = ""
TextBox48.Text = ""
TextBox43.Text = ""
TextBox45.Text = ""
TextBox46.Text = ""
TextBox47.Text = ""
TextBox52.Text = ""
TextBox49.Text = ""
TextBox50.Text = ""
TextBox51.Text = ""

TextBox21.Text = ""
TextBox22.Text = ""


Unload Me
End Sub

Private Sub CommandButton15_Click()
DISTRITOS_3.Show
End Sub

Private Sub CommandButton16_Click()
OCUPACION2.Show
End Sub

Private Sub CommandButton17_Click()
DISTRITOS_4.Show
End Sub

Private Sub CommandButton18_Click()
OCUPACION3.Show
End Sub

Private Sub CommandButton19_Click()
DISTRITOS_5.Show
End Sub

Private Sub CommandButton20_Click()
OCUPACION4.Show
End Sub

Private Sub TextBox21_Change()
Sheets("LAVA").Select

Cells(43, 2) = TextBox21.Text


End Sub

Private Sub TextBox21_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox21.Text = UCase(TextBox21.Text)
End Sub

Private Sub TextBox22_Change()
Sheets("LAVA").Select
Cells(47, 2) = TextBox22.Text
End Sub

Private Sub TextBox22_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox22.Text = UCase(TextBox22.Text)
End Sub

Private Sub TextBox23_Change()
Sheets("LAVA").Select
Cells(18, 3) = TextBox23.Text
End Sub

Private Sub TextBox23_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox23.Text = UCase(TextBox23.Text)

End Sub

Private Sub TextBox24_Change()
Sheets("LAVA").Select
TextBox24.MaxLength = 8
Cells(16, 5) = TextBox24.Text
End Sub

Private Sub TextBox24_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox28_Change()
Sheets("LAVA").Select
Cells(17, 3) = TextBox28.Text
End Sub

Private Sub TextBox28_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox28.Text = UCase(TextBox28.Text)

End Sub

Private Sub TextBox29_Change()

TextBox29.MaxLength = 10
largo_entrada = Len(Me.TextBox29)
Select Case largo_entrada
Case 2
Me.TextBox29.Value = Me.TextBox29.Value & "/"
Case 5
Me.TextBox29.Value = Me.TextBox29.Value & "/"

End Select



Sheets("LAVA").Select
Cells(20, 6) = Format(TextBox29.Text, "mm/dd/yyyy")
End Sub

Private Sub TextBox30_Change()
Sheets("LAVA").Select
TextBox30.MaxLength = 9
Cells(20, 9) = TextBox30.Text
End Sub

Private Sub TextBox31_Change()
Sheets("LAVA").Select
Cells(21, 3) = TextBox31.Text
End Sub

Private Sub TextBox32_Change()
Sheets("LAVA").Select
Cells(20, 3) = TextBox32.Text
End Sub

Private Sub TextBox32_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox32.Text = UCase(TextBox32.Text)
End Sub

Private Sub TextBox33_Change()
Sheets("LAVA").Select
Cells(28, 3) = TextBox33.Text

End Sub

Private Sub TextBox33_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox33.Text = UCase(TextBox33.Text)
End Sub

Private Sub TextBox34_Change()
Sheets("LAVA").Select
TextBox34.MaxLength = 11
Cells(26, 5) = TextBox34.Text

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
Sheets("LAVA").Select
Cells(29, 3) = TextBox35.Text

End Sub

Private Sub TextBox36_Change()
Sheets("LAVA").Select
Cells(29, 6) = TextBox36.Text

End Sub

Private Sub TextBox37_Change()
Sheets("LAVA").Select
Cells(29, 9) = TextBox37.Text

End Sub

Private Sub TextBox38_Change()
Sheets("LAVA").Select
Cells(27, 3) = TextBox38.Text

End Sub

Private Sub TextBox38_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox38.Text = UCase(TextBox38.Text)
End Sub

Private Sub TextBox39_Change()
TextBox39.MaxLength = 10
largo_entrada = Len(Me.TextBox39)
Select Case largo_entrada
Case 2
Me.TextBox39.Value = Me.TextBox39.Value & "/"
Case 5
Me.TextBox39.Value = Me.TextBox39.Value & "/"

End Select



Sheets("LAVA").Select
Cells(30, 6) = Format(TextBox39.Text, "mm/dd/yyyy")



End Sub

Private Sub TextBox40_Change()
Sheets("LAVA").Select
TextBox40.MaxLength = 9
Cells(30, 9) = TextBox40.Text

End Sub

Private Sub TextBox41_Change()
Sheets("LAVA").Select
Cells(31, 3) = TextBox41.Text
End Sub

Private Sub TextBox42_Change()
Sheets("LAVA").Select

Cells(30, 3) = TextBox42.Text

End Sub

Private Sub TextBox42_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox42.Text = UCase(TextBox42.Text)
End Sub

Private Sub TextBox43_Change()
Sheets("LAVA").Select

Cells(37, 3) = TextBox43.Text

End Sub

Private Sub TextBox43_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox43.Text = UCase(TextBox43.Text)
End Sub

Private Sub TextBox44_Change()
Sheets("LAVA").Select
TextBox44.MaxLength = 11
Cells(35, 5) = TextBox44.Text


End Sub

Private Sub TextBox44_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox45_Change()
Sheets("LAVA").Select

Cells(38, 3) = TextBox45.Text

End Sub

Private Sub TextBox46_Change()
Sheets("LAVA").Select

Cells(38, 6) = TextBox46.Text
End Sub

Private Sub TextBox47_Change()
Sheets("LAVA").Select

Cells(38, 9) = TextBox47.Text


End Sub

Private Sub TextBox48_Change()
Sheets("LAVA").Select

Cells(36, 3) = TextBox48.Text


End Sub

Private Sub TextBox48_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox48.Text = UCase(TextBox48.Text)
End Sub

Private Sub TextBox49_Change()
TextBox49.MaxLength = 10
largo_entrada = Len(Me.TextBox49)
Select Case largo_entrada
Case 2
Me.TextBox49.Value = Me.TextBox49.Value & "/"
Case 5
Me.TextBox49.Value = Me.TextBox49.Value & "/"

End Select



Sheets("LAVA").Select
Cells(39, 6) = Format(TextBox49.Text, "mm/dd/yyyy")


End Sub

Private Sub TextBox50_Change()
Sheets("LAVA").Select
TextBox50.MaxLength = 9
Cells(39, 9) = TextBox50.Text

End Sub

Private Sub TextBox51_Change()
Sheets("LAVA").Select

Cells(40, 3) = TextBox51.Text

End Sub

Private Sub TextBox52_Change()
Sheets("LAVA").Select

Cells(39, 3) = TextBox52.Text


End Sub

Private Sub TextBox52_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox52.Text = UCase(TextBox52.Text)
End Sub

Private Sub UserForm_activate()
ComboBox3.AddItem ("DNI")
ComboBox3.AddItem ("CE")
ComboBox4.AddItem ("DNI")
ComboBox4.AddItem ("CE")
ComboBox4.AddItem ("RUC")
ComboBox5.AddItem ("DNI")
ComboBox5.AddItem ("CE")
ComboBox5.AddItem ("RUC")
Sheets("LAVA").Select
Cells(10, 3) = MENU.TextBox1.Text
Cells(11, 3) = MENU.TextBox4.Text
botonciiu2.Visible = False
botonciiu3.Visible = False

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.ScrollBars = fmScrollBarsVertical ' Esto habilita la barra de desplazamiento vertical
    Application.Visible = False ' Oculta el libro de Excel
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True ' Vuelve a mostrar el libro de Excel al cerrar el UserForm
End Sub



