VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PASE 
   Caption         =   " - SIAF"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445.001
   OleObjectBlob   =   "PASE.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "PASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
                Cells(3, 4) = ComboBox5.Text
                Cells(3, 5) = ComboBox4.Text
                Cells(3, 3) = "Pago de Servicio"
                Cells(3, 6) = "Efectivo"
                Cells(3, 7) = TextBox3.Text
                Cells(3, 8) = ""
                Cells(3, 9) = TextBox1.Value
                Cells(3, 10) = ""
                Cells(3, 11) = ""
                Cells(3, 12) = ""

                      
                Else
                Sheets("ULTIMO REGISTRO").Select
                Cells(3, 2) = Label8.Caption
                Cells(3, 4) = ComboBox5.Text
                Cells(3, 5) = ComboBox4.Text
                Cells(3, 3) = "Pago de servicio"
                Cells(3, 6) = "Efectivo"
                Cells(3, 7) = TextBox3.Text
                Cells(3, 8) = ""
                Cells(3, 9) = ""
                Cells(3, 10) = ""
                Cells(3, 11) = TextBox1.Value
                Cells(3, 12) = ""

       
                End If
               'Codigo obtenido del grabador de macros
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
                MsgBox " Registrado Correctamente", vbInformation, "SIAF"
                 VREDE.Show
                ComboBox4.Text = ""
                                TextBox3.Text = ""
                TextBox1.Value = ""
                Frame1.BackColor = &H80000010
                Label5.BackColor = &H80000010
                Label6.BackColor = &H80000010
                Label7.BackColor = &H80000010
                 Me.Hide
                End If
                          
   

End Sub

Private Sub CommandButton2_Click()
    ComboBox4.Text = ""
    TextBox3.Text = ""
    TextBox1.Value = ""
    Frame1.BackColor = &H80000010
    Label5.BackColor = &H80000010
    Label6.BackColor = &H80000010
    Label7.BackColor = &H80000010
    Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = False
Sheets("ULTIMO REGISTRO").Visible = False
Sheets("TIPO DE CAMBIO").Visible = False
Sheets("ULTIMA CUENTA").Visible = False
Sheets("BASE CUENTAS").Visible = False

    Me.Hide
End Sub

Private Sub TextBox1_AfterUpdate()
TextBox1 = Format(TextBox1.Value, "#,###,###,##0.00")
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_activate()
      
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Label8.Caption = TimeValue(Now)
Application.ScreenUpdating = False
Application.Visible = False
End Sub


Private Sub UserForm_Initialize()
ComboBox4.AddItem ("MN S/")
      ComboBox4.AddItem ("ME $")
      ComboBox5.AddItem ("SERVICIO LUZ")
      ComboBox5.AddItem ("SERVICIO DE AGUA")
      ComboBox5.AddItem ("SERVICIO DE TELEFONÍA Y CABLE")
      ComboBox5.AddItem ("SERVICIO DE INTERNET")
      ComboBox5.AddItem ("SERVICIO DE TELEFONÍA MOVIL")
      ComboBox5.AddItem ("SERVICIO DE SEGURIDAD ELECTRÓNICA")
      ComboBox5.AddItem ("CERTUS")
      ComboBox5.AddItem ("UNIVERSIDAD PRIVADA")
      ComboBox5.AddItem ("UNIVERSIDAD PÚBLICA")
      ComboBox5.AddItem ("COLEGIO PARTICULAR")
   ComboBox5.AddItem ("SUNAT")
   ComboBox5.AddItem ("ONP")
   ComboBox5.AddItem ("PRIMA AFP")
   ComboBox5.AddItem ("AFP INTEGRA")
End Sub


