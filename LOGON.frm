VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LOGON 
   Caption         =   "DOMAIN LOGON - SIAF"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   OleObjectBlob   =   "LOGON.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "LOGON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()


    
        If TextBox1 = "admin" And TextBox2 = "admin" Then
            TextBox1.Text = ""
            TextBox1.Text = ""
            
            Dim Resp As Byte
                Resp = MsgBox("¿Deseas generar nuevo registro?", _
                vbQuestion + vbYesNo, "SIAF")
            If Resp = vbYes Then
                MsgBox "SIAF sestá generando un nuevo reporte diario, espere un momento por favor...", vbExclamation, "SIAF"
                Sheets("REPORTE MONETARIO").Visible = True
                Sheets("REPORTE MONETARIO").Select
                ActiveSheet.Unprotect
                Range("B1:B4,D3:D4,E1:E2,A9:L241").Select
                Range("A9").Activate
                Selection.ClearContents
                ActiveWindow.SmallScroll Down:=-9
                TextBox1.Text = ""
                TextBox1.Text = ""
                Me.Hide
                INGRESO.Show
            Else
                MsgBox "SIAF sestá cargando el reporte diario anterior, espere un momento por favor...", vbExclamation, "SIAF"
                Sheets("REPORTE MONETARIO").Visible = True
                Sheets("REPORTE MONETARIO").Select
                ActiveSheet.Unprotect
                ActiveWindow.SmallScroll Down:=-9
                
                Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
                Sheets("ULTIMO REGISTRO").Visible = True
                Sheets("TIPO DE CAMBIO").Visible = True
                Sheets("ULTIMA CUENTA").Visible = True
                Sheets("BASE CUENTAS").Visible = True
                
                Unload Me
                
                MENU.Show
            End If
       Else
            MsgBox ("Usuario/Contraseña incorrectos"), vbInformation
            TextBox1.Text = ""
            TextBox2 = ""
      
        Me.Hide
        TextBox1.Text = ""
        TextBox1.Text = ""
   
        
End If

End Sub

Private Sub CommandButton2_Click()
Me.Hide
Application.Visible = True
Sheets("INICIO").Select
End Sub
Private Sub Userform1_Initialize()
lbl_Fecha.Caption = Date
lbl_Hora.Caption = Time
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_activate()
Application.ScreenUpdating = False
Label8.Caption = TimeValue(Now)
Application.ScreenUpdating = False
Application.Visible = False

End Sub

Private Sub UserForm_Terminate()
Me.Hide
Application.Visible = True
Sheets("INICIO").Select
End Sub
