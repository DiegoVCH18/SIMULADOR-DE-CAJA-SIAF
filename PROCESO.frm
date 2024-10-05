VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PROCESO 
   Caption         =   "*** PROCESO ***"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12675
   OleObjectBlob   =   "PROCESO.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "PROCESO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
MsgBox "SIAF se está calculando las cuotas, espere un momento por favor...", vbExclamation, "SIAF"
SOLICITUDCP.Frame3.Visible = True
SOLICITUDCP.TextBox13.Text = Sheets("SOLICITUD CP").Cells(84, 15)
SOLICITUDCP.TextBox14.Text = Format(Sheets("SOLICITUD CP").Cells(80, 12), "0.00%")
SOLICITUDCP.TextBox15.Text = Format(Sheets("SOLICITUD CP").Cells(128, 7), "0.00%")
SOLICITUDCP.TextBox20.Text = Sheets("SOLICITUD CP").Cells(147, 7)
SOLICITUDCP.TextBox16.Text = Sheets("SOLICITUD CP").Cells(229, 5)
SOLICITUDCP.TextBox17.Text = Sheets("SOLICITUD CP").Cells(229, 7)
SOLICITUDCP.TextBox18.Text = Sheets("SOLICITUD CP").Cells(229, 11)
SOLICITUDCP.TextBox19.Text = Sheets("SOLICITUD CP").Cells(229, 12)
SOLICITUDCP.TextBox21.Text = Sheets("SOLICITUD CP").Cells(103, 4)
SOLICITUDCP.TextBox8.Text = Sheets("SOLICITUD CP").Cells(103, 12)
Me.Hide
End Sub
