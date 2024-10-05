VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LAVA2 
   Caption         =   "REGISTRO DE OPERACIONES EN EFECTIVO - SIAF"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5025
   OleObjectBlob   =   "LAVA2.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "LAVA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton14_Click()
TextBox24.Text = ""
ComboBox3.Text = ""
DEPO.TextBox15 = "COMPLETO"
                        DEPO.TextBox15 = "COMPLETO"
                        CANC.TextBox15 = "COMPLETO"
                        RETI.TextBox15 = "COMPLETO"
                        CHPA.TextBox15 = "COMPLETO"
                        COBR.TextBox15 = "COMPLETO"
                        EMIS.TextBox15 = "COMPLETO"
                        PAGO.TextBox15 = "COMPLETO"
                        Unload Me
End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub UserForm_activate()
ComboBox3.AddItem ("DNI")
ComboBox3.AddItem ("CE")
End Sub

Private Sub UserForm_Click()

End Sub

