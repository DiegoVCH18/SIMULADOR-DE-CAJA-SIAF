VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OCUPACION2 
   Caption         =   "OCUPACIÓN CLIENTE - SIAF"
   ClientHeight    =   8565.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "OCUPACION2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "OCUPACION2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Sheets("DATOS GENERALES").Select
LAVA.TextBox31.Value = Cells(488, 5)

Unload Me
End Sub

Private Sub LISTA_Click()
Dim CODIGO As Integer
CODIGO = LISTA.List(LISTA.ListIndex, 0)
Me.txt_codigo.Text = CODIGO

Sheets("DATOS GENERALES").Select
Cells(487, 5) = CODIGO
TextBox1.Text = Cells(488, 5).Text

End Sub

Private Sub TEXTO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub



Private Sub TextBox1_Change()

End Sub

Private Sub Txt_busqueda_Change(): On Error Resume Next
Dim fila, Final, i As Long
fila = 2
Do While Sheets("DATOS GENERALES").Cells(fila, 16) <> Empty
fila = fila + 1
Loop
Final = fila - 1
LISTA.Clear

For i = 2 To Final
       
    If UCase(Sheets("DATOS GENERALES").Cells(i, 16)) Like "*" & UCase(txt_busqueda) & "*" Then
       With LISTA
       .AddItem
       .List(.ListCount - 1, 0) = Sheets("DATOS GENERALES").Cells(i, 15)
       .List(.ListCount - 1, 1) = Sheets("DATOS GENERALES").Cells(i, 16)
       End With
     End If
     
Next i

End Sub
Private Sub txt_busqueda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 250) Then
KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub
Private Sub txt_codigo_Change()

End Sub

Private Sub UserForm_Initialize(): On Error Resume Next

Dim fila, Final, i As Long
fila = 2
Do While Sheets("DATOS GENERALES").Cells(fila, 1) <> Empty
fila = fila + 1
Loop
Final = fila - 1

For i = 2 To Final

       With LISTA
            .AddItem
            .List(.ListCount - 1, 0) = Sheets("DATOS GENERALES").Cells(i, 15)
            .List(.ListCount - 1, 1) = Sheets("DATOS GENERALES").Cells(i, 16)
            
       End With
          
Next i

       With LISTA
       .ColumnCount = 2
       .ColumnWidths = "40 pt;100 pt"
       End With
       
End Sub



