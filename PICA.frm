VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PICA 
   Caption         =   "PROVISIÓN INICIAL DE CAJA  - SIAF"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   OleObjectBlob   =   "PICA.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "PICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'propiedad intelectual ing. Diego Armando Vásquez Chávez DNI: 44113825
Private Sub ComboBox1_Change()
Dim a As Double
Dim b As Double
Sheets("REPORTE MONETARIO").Select
a = Cells(1, 9)
b = Cells(2, 9)
If ComboBox1.Text = "MN S/" Then
Label7.Caption = "S/"
Label8.Caption = "S/"
Label53.Caption = "S/"
TextBox1.Visible = True
Label6.Visible = True
Label9.Visible = True
Label5.Visible = True
Label10.Visible = True
Else
Label7.Caption = "US$"
Label8.Caption = "US$"
Label53.Caption = "US$"
TextBox1.Visible = False
Label6.Visible = False
Label9.Visible = False
Label5.Visible = False
Label10.Visible = False
End If
End Sub

Private Sub CommandButton1_Click()


If ComboBox1.Text = "MN S/" Then
    If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
    MsgBox ("Completar todas las casillas"), , "SIAF v 1.3.0"
    Else
    Label10.Caption = Val(Label6.Caption * TextBox1.Value)
    Label11.Caption = Val(Label13.Caption * TextBox2.Value)
    Label15.Caption = Val(Label17.Caption * TextBox3.Value)
    Label19.Caption = Val(Label21.Caption * TextBox4.Value)
    Label23.Caption = Val(Label25.Caption * TextBox5.Value)
    
    TextBox13.Text = Val(Val(Label10.Caption) + Val(Label11.Caption) + Val(Label15.Caption) + Val(Label19.Caption) + Val(Label23.Caption))
    TextBox13 = Format(TextBox13.Value, "#,###,###,##0.00")
    End If
Else
    If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
    MsgBox ("Completar todas las casillas"), , "SIAF v 1.3.0"
    Else
    Label11.Caption = Val(Label13.Caption * TextBox2.Value)
    Label15.Caption = Val(Label17.Caption * TextBox3.Value)
    Label19.Caption = Val(Label21.Caption * TextBox4.Value)
    Label23.Caption = Val(Label25.Caption * TextBox5.Value)
    TextBox13.Text = Val(Val(Label11.Caption) + Val(Label15.Caption) + Val(Label19.Caption) + Val(Label23.Caption))
    TextBox13 = Format(TextBox13.Value, "#,###,###,##0.00")
    End If
End If
End Sub

Private Sub CommandButton2_Click()
If TextBox13.Text = "" Then
MsgBox "falta totalizar", , "SIAF v 1.3.0"
Else
If ComboBox1.Text = "MN S/" Then
Else
    End If
     If ComboBox1.Text = "MN S/" Then
    Sheets("ULTIMO REGISTRO").Select
                Cells(3, 2) = Label61.Caption
                Cells(3, 4) = "Interno"
                Cells(3, 5) = ComboBox1.Text
                Cells(3, 3) = "Recepción dinero boveda"
                Cells(3, 6) = "Efectivo"
                Cells(3, 7) = "-"
                Cells(3, 8) = "-"
                Cells(3, 9) = TextBox13.Value
                Cells(3, 10) = "-"
                Cells(3, 11) = "-"
                Cells(3, 12) = "-"
Sheets("REPORTE MONETARIO").Select
                Cells(3, 4) = TextBox13.Value
                                      
                Else
                Sheets("ULTIMO REGISTRO").Select
                Cells(3, 2) = Label61.Caption
                Cells(3, 4) = "Interno"
                Cells(3, 5) = ComboBox1.Text
                Cells(3, 3) = "Recepción dinero boveda"
                Cells(3, 6) = "Efectivo"
                Cells(3, 7) = "-"
                Cells(3, 8) = "-"
                Cells(3, 9) = "-"
                Cells(3, 10) = "-"
                Cells(3, 11) = TextBox13.Value
                Cells(3, 12) = "-"
Sheets("REPORTE MONETARIO").Select
                Cells(4, 4) = TextBox13.Value
       
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
                MsgBox " Registrado Correctamente", vbInformation, "SIAF v 1.3.0"
                VREDE.Show
                ComboBox1.Text = ""
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox13.Text = ""
    Label10.Caption = ""
    Label11.Caption = ""
    Label15.Caption = ""
    Label19.Caption = ""
    Label23.Caption = ""
    MENU.Label9.Visible = True
    Me.Hide
    CONSULTA.Show
    End If
End Sub

Private Sub CommandButton3_Click()
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = False
Sheets("ULTIMO REGISTRO").Visible = False
Sheets("TIPO DE CAMBIO").Visible = False
Sheets("ULTIMA CUENTA").Visible = False
Sheets("BASE CUENTAS").Visible = False

ComboBox1.Text = ""
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox13.Text = ""
    Label10.Caption = ""
    Label11.Caption = ""
    Label15.Caption = ""
    Label19.Caption = ""
    Label23.Caption = ""
Me.Hide
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub TextBox1_Keypress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub



Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
      'Código para restringir ingreso a solo numeros
      If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
      status_bar = "!!!ATENCION!!!, este campo es únicamente numérico."
      Else
      'no pasa nada
      End If
End Sub

Private Sub UserForm_activate()

   
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Label61.Caption = TimeValue(Now)
Application.ScreenUpdating = False
Application.Visible = False
Label10.Caption = "0.00"
    Label11.Caption = "0.00"
    Label15.Caption = "0.00"
    Label19.Caption = "0.00"
    Label23.Caption = "0.00"
End Sub



Private Sub UserForm_Initialize()
ComboBox1.AddItem ("MN S/")
   ComboBox1.AddItem ("US $")
   Label10.Caption = "0.00"
    Label11.Caption = "0.00"
    Label15.Caption = "0.00"
    Label19.Caption = "0.00"
    Label23.Caption = "0.00"
End Sub


