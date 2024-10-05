VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SALIDA 
   Caption         =   "Desconexi�n"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   OleObjectBlob   =   "SALIDA.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "SALIDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub CommandButton1_Click()

Application.ScreenUpdating = False
Application.Visible = True
Sheets("REPORTE MONETARIO").Visible = True
Sheets("REPORTE MONETARIO").Select
ActiveWindow.DisplayHeadings = False
Cells.Select
ActiveSheet.Protect
ActiveWindow.Zoom = 150
ActiveWindow.DisplayHeadings = False
ExecuteExcel4Macro ("show.toolbar(""ribbon"",0)")
ActiveWindow.SmallScroll Down:=-15
ActiveWindow.DisplayHorizontalScrollBar = False
Sheets("CARACTER�STICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Dim Resp As Byte
Resp = MsgBox("�Deseas salir?", _
    vbQuestion + vbYesNo, "SIAF")
If Resp = vbYes Then
    MsgBox "El SIAF se est� cerrando, espere un momento por favor...", vbExclamation, "SIAF"
    Sheets("REPORTE MONETARIO").Visible = True
    Sheets("CARACTER�STICAS OPERATIVAS").Visible = True
    Sheets("ULTIMO REGISTRO").Visible = True
    Sheets("TIPO DE CAMBIO").Visible = True
    Sheets("ULTIMA CUENTA").Visible = True
    Sheets("BASE CUENTAS").Visible = True
   
    Sheets("INICIO").Visible = True
    ThisWorkbook.Save
    MsgBox "Gracias por utilizar SIAF", vbExclamation, "SIAF"
    ThisWorkbook.Close
    
Else
    MsgBox "Se eligi� cancelar...", vbCritical, "SIAF"
    MENU.Show
    
    
End If

End Sub

Private Sub UserForm_Initialize()
Application.ScreenUpdating = False
Application.Visible = True
End Sub

