VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PROCESO2 
   Caption         =   "***PROCESO***"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   OleObjectBlob   =   "PROCESO2.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "PROCESO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

MsgBox "SIAF se está procesando la solicitud, espere un momento por favor...", vbExclamation, "SIAF"
SOLICITUDTC.TextBox13.Text = Sheets("SOLICITUD TC").Cells(37, 6)
SOLICITUDTC.TextBox14.Text = Sheets("SOLICITUD TC").Cells(37, 10)
SOLICITUDTC.Frame3.Visible = True
Me.Hide
End Sub
