VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INGRESO 
   Caption         =   "PROCESO CORRECTO"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8145
   OleObjectBlob   =   "INGRESO.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "INGRESO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton2_Click()
Sheets("REPORTE MONETARIO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Me.Hide
CONEXIÓN.Show
End Sub

Private Sub UserForm_Initialize()
Application.ScreenUpdating = False
End Sub

Private Sub UserForm_Terminate()
Me.Hide
LOGON.Show
End Sub
