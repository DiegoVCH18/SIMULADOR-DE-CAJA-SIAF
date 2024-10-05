VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PAGO 
   Caption         =   "PAGO  - SIAF"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   OleObjectBlob   =   "PAGO.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "PAGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

PRODUCTOS.Show

End Sub

Private Sub ComboBox1_Change()
If ComboBox1.Text = "1.TARJETA DE CRÉDITO" Then
TextBox7.MaxLength = 19
Label3.Caption = "N° de tarjeta:"
Else
TextBox7.MaxLength = 18
Label3.Caption = "N° de crédito:"

End If
End Sub

Private Sub ComboBox4_Change()
Label7.Caption = ComboBox4.Text


If ComboBox4.Text = "MN S/" Then

Frame1.BackColor = &H80000003
Label5.BackColor = &H80000003
Label6.BackColor = &H80000003
Label7.BackColor = &H80000003





Else

Frame1.BackColor = &H80FF80
Label5.BackColor = &H80FF80
Label6.BackColor = &H80FF80
Label7.BackColor = &H80FF80

End If

End Sub

Private Sub CommandButton1_Click()
If TextBox1 = "" Then
    MsgBox "Ingresar Cantidad ", vbInformation, "SIAF"
    Else
            If ComboBox4.Text = "MN S/" Then
                Sheets("ULTIMO REGISTRO").Select
                Cells(3, 2) = Label8.Caption
                Cells(3, 4) = ComboBox1.Text
                Cells(3, 5) = ComboBox4.Text
                Cells(3, 3) = "Depósito"
                Cells(3, 6) = "Efectivo"
                Cells(3, 7) = TextBox7.Value
                Cells(3, 8) = ""
                Cells(3, 9) = TextBox1.Value
                Cells(3, 10) = ""
                Cells(3, 11) = ""
                Cells(3, 12) = ""

                Sheets("LAVA").Select
                TextBox13.Text = Cells(54, 14)
                TextBox14.Text = Cells(55, 14)

                            If TextBox13.Text = "LAVA" Then
                               LAVA.Show
                            Else
                                If TextBox13.Text = "DNI" Then
                                    LAVA2.Show
                                End If
                            End If
                            If TextBox13.Text = "NADA" Then
                            TextBox15.Text = "COMPLETO"
                    End If
            Else
                Sheets("ULTIMO REGISTRO").Select
                Cells(3, 2) = Label8.Caption
                Cells(3, 4) = ComboBox1.Text
                Cells(3, 5) = ComboBox4.Text
                Cells(3, 3) = "Pago"
                Cells(3, 6) = "Efectivo"
                Cells(3, 7) = TextBox7.Value
                Cells(3, 8) = ""
                Cells(3, 9) = ""
                Cells(3, 10) = ""
                Cells(3, 11) = TextBox1.Value
                Cells(3, 12) = ""

                Sheets("LAVA").Select
                TextBox13.Text = Cells(54, 14)
                TextBox14.Text = Cells(55, 14)
                   If TextBox14.Text = "DNI" Then
                        LAVA2.Show
                    Else
                        If TextBox14.Text = "LAVA" Then
                        LAVA.Show
                        End If
                    End If
                    If TextBox13.Text = "NADA" Then
                    TextBox15.Text = "COMPLETO"
                    End If
            End If
                
                If TextBox14.Text = "DNI" Then
                        If TextBox15.Text = "COMPLETO" Then
                           'Codigo obtenido del grabador de macros
                            Application.Visible = True
                           Sheets("REPORTE MONETARIO").Select
                           Rows("9:9").Select
                           Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                           With Selection.Interior
                               .Pattern = xlNone
                               .TintAndShade = 0
                               .PatternTintAndShade = 0
                           End With
                           Sheets("ULTIMO REGISTRO").Select
                           Range("A3:O3").Select
                           Selection.Copy
                           Range("A3").Select
                           Sheets("REPORTE MONETARIO").Select
                           Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
                           Sheets("REPORTE MONETARIO").Select
                           MsgBox " Registrado Correctamente", vbExclamation, "SIAF"
                               ComboBox1.Text = ""
                            ComboBox4.Text = ""
                            TextBox7.Text = ""
                            TextBox1.Text = ""
                            Frame1.BackColor = &H80000010
                           Label5.BackColor = &H80000010
                           Label6.BackColor = &H80000010
                           Label7.BackColor = &H80000010
                           Unload Me
                        Else
                            MsgBox "Completar Registro de operaciones en efectivo de mayor cuantía", vbCritical, "SIAF"
                        End If
                    Else
                     If TextBox14.Text = "LAVA" Then
                     
                        If TextBox15.Text = "COMPLETO" Then
                        
                          'Codigo obtenido del grabador de macros
                          Application.Visible = True
                           Sheets("REPORTE MONETARIO").Select
                           Rows("9:9").Select
                           Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                           With Selection.Interior
                               .Pattern = xlNone
                               .TintAndShade = 0
                               .PatternTintAndShade = 0
                           End With
                           Sheets("ULTIMO REGISTRO").Select
                           Range("A3:O3").Select
                           Selection.Copy
                           Range("A3").Select
                           Sheets("REPORTE MONETARIO").Select
                           Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
                           Sheets("REPORTE MONETARIO").Select
                           MsgBox " Registrado Correctamente", vbExclamation, "SIAF"
                               ComboBox1.Text = ""
                            ComboBox4.Text = ""
                            TextBox7.Text = ""
                            TextBox1.Text = ""
                            Frame1.BackColor = &H80000010
                           Label5.BackColor = &H80000010
                           Label6.BackColor = &H80000010
                           Label7.BackColor = &H80000010
                           Unload Me
                        Else
                        MsgBox "Completar Registro de operaciones en efectivo", vbCritical, "SIAF"
                        End If
                    Else
                    'Codigo obtenido del grabador de macros
                    If TextBox14.Text = "NADA" Then
                          Application.Visible = True
                           Sheets("REPORTE MONETARIO").Select
                           Rows("9:9").Select
                           Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                           With Selection.Interior
                               .Pattern = xlNone
                               .TintAndShade = 0
                               .PatternTintAndShade = 0
                           End With
                           Sheets("ULTIMO REGISTRO").Select
                           Range("A3:O3").Select
                           Selection.Copy
                           Range("A3").Select
                           Sheets("REPORTE MONETARIO").Select
                           Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
                           Sheets("REPORTE MONETARIO").Select
                           MsgBox " Registrado Correctamente", vbExclamation, "SIAF"
                               VREDE.Show
                           ComboBox1.Text = ""
                            ComboBox4.Text = ""
                            TextBox7.Text = ""
                            TextBox1.Text = ""
                            Frame1.BackColor = &H80000010
                           Label5.BackColor = &H80000010
                           Label6.BackColor = &H80000010
                           Label7.BackColor = &H80000010
                           Unload Me
                    End If
               
        
             End If

                End If
             
   End If
    Unload CHPA
    Unload RETI
    Unload DEPO
    Unload EMIS
    Unload CANC
    Unload COBR
End Sub

Private Sub CommandButton2_Click()
ComboBox1.Text = ""
             ComboBox4.Text = ""
             TextBox7.Text = ""
             TextBox1.Text = ""
             
             
             Frame1.BackColor = &H80000010
                Label5.BackColor = &H80000010
                Label6.BackColor = &H80000010
                Label7.BackColor = &H80000010
                Me.Hide
                Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = False
Sheets("ULTIMO REGISTRO").Visible = False
Sheets("TIPO DE CAMBIO").Visible = False
Sheets("ULTIMA CUENTA").Visible = False
Sheets("BASE CUENTAS").Visible = False
DEPO.TextBox12.Text = ""
DEPO.ComboBox1.Text = ""
DEPO.ComboBox4.Text = ""
DEPO.TextBox7.Text = ""
RETI.TextBox12.Text = ""
RETI.ComboBox1.Text = ""
RETI.ComboBox4.Text = ""
RETI.TextBox7.Text = ""
CANC.TextBox12.Text = ""
CANC.ComboBox1.Text = ""
CANC.ComboBox4.Text = ""
CANC.TextBox7.Text = ""
PAGO.
End Sub

Private Sub OptionButton1_Click()
TextBox12.Visible = True
End Sub



Private Sub TextBox12_Change()
Dim nro As String
Dim cta As String
Dim tipo As String
Dim Moneda As String
ult = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To ult
    Sheets("BASE CUENTAS").Select
    nro = Cells(i, 1)
    cta = Cells(i, 7)
    tipo = Cells(i, 5)
    Moneda = Cells(i, 6)
    If TextBox12.Text = nro Then
        TextBox7.Text = cta
        ComboBox1.Text = tipo
        ComboBox4.Text = Moneda
    End If
Next
End Sub

Private Sub TextBox9_Change()
TextBox11.Text = Val(TextBox1.Text) + Val(TextBox9.Text)
End Sub

Private Sub TextBox1_AfterUpdate()
TextBox1 = Format(TextBox1.Value, "#,###,###,##0.00")
End Sub

Private Sub TextBox1_Change()
Sheets("LAVA").Select
TextBox13.Text = Cells(54, 14)
TextBox14.Text = Cells(55, 14)

                    
End Sub

Private Sub TextBox7_Change()
If ComboBox1.Text = "1.TARJETA DE CRÉDITO" Then
TextBox7.MaxLength = 19
largo_entrada = Len(Me.TextBox7)
Select Case largo_entrada
    Case 4
        Me.TextBox7.Value = Me.TextBox7.Value & "-"
    Case 9
        Me.TextBox7.Value = Me.TextBox7.Value & "-"
    Case 14
        Me.TextBox7.Value = Me.TextBox7.Value & "-"
End Select
Else
TextBox7.MaxLength = 21
largo_entrada = Len(Me.TextBox7)
Select Case largo_entrada
    Case 3
        Me.TextBox7.Value = Me.TextBox7.Value & "-"
    Case 7
        Me.TextBox7.Value = Me.TextBox7.Value & "-"
End Select
Sheets("datos generales").Select
Cells(1, 9) = TextBox7.Text
If Cells(1, 10) = "101" Then
ComboBox4.Text = "ME $"
Else
ComboBox4.Text = "MN S/"
End If
End If
End Sub

Private Sub UserForm_activate()

Label8.Caption = TimeValue(Now)
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Sheets("datos generales").Select
Cells(5, 9) = ""

Application.ScreenUpdating = False
Application.Visible = False
End Sub

Private Sub UserForm_Initialize()
ComboBox4.AddItem ("MN S/")
ComboBox4.AddItem ("ME $")
ComboBox1.AddItem ("1.TARJETA DE CRÉDITO")
ComboBox1.AddItem ("2.CRÉDITO CONSUMO")
ComboBox1.AddItem ("3.CRÉDITO PYME")
ComboBox1.AddItem ("4.CRÉDITO VEHICULAR")
ComboBox1.AddItem ("5.CRÉDITO HIPOTECARIO")

End Sub


