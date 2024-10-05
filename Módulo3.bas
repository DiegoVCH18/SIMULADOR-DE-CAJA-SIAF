Attribute VB_Name = "Módulo3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("REPORTE MONETARIO").Select
    Range("B1:B4,A9:L116").Select
    Range("A9").Activate
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("B1:B4").Select
    Sheets("INICIO").Select
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("REPORTE MONETARIO").Select
    Range("B1:B4,E1:E2,A9:L241").Select
    Range("A9").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-9
    Range("A9").Select
    Sheets("INICIO").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("A1:L161").Select
    Range("K1").Activate
End Sub
Sub Botón3_Haga_clic_en()
Dim Resp As Byte
Resp = MsgBox("¿Deseas salir?", _
    vbQuestion + vbYesNo, "EXCELeINFO")
If Resp = vbYes Then
    MsgBox "El SIAF se está cerrando, espere un momento por favor...", vbExclamation, "EXCELeINFO"
    Sheets("REPORTE MONETARIO").Visible = False
    ThisWorkbook.Save
    ThisWorkbook.Close
Else
    MsgBox "Se eligió cancelar...", vbCritical, "EXCELeINFO"
End If
End Sub
