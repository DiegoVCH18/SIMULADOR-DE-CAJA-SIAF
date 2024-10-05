VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MENU 
   Caption         =   "MENU DE TRANSACCIONES"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16935
   OleObjectBlob   =   "MENU.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
RETI.Show

End Sub
Private Sub CommandButton11_Click()
EMIS.Show

End Sub
Private Sub CommandButton12_Click()
COBR.Show

End Sub
Private Sub CommandButton13_Click()
Me.Hide
PICA.Show


End Sub
Private Sub CommandButton14_Click()
DIRE.Show

End Sub
Private Sub CommandButton15_Click()
DIEN.Show

End Sub
Private Sub CommandButton16_Click()
COME.Show

End Sub
Private Sub CommandButton17_Click()
VEME.Show

End Sub
Private Sub CommandButton19_Click()
Me.Hide
Worksheets("REPORTE MONETARIO").PrintPreview
End Sub
Private Sub CommandButton2_Click()
DEPO.Show

End Sub
Private Sub CommandButton20_Click()
SALIDA.Show

End Sub
Private Sub CommandButton21_Click()
APERTURA.Show

End Sub
Private Sub CommandButton22_Click()
CCFI.Show

End Sub
Private Sub CommandButton23_Click()
Application.ScreenUpdating = False
Me.Hide
CONSULTA.Show
End Sub
Private Sub CommandButton24_Click()
CONSTICA.Show

End Sub
Private Sub CommandButton29_Click()
Application.ScreenUpdating = False
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("TIPO DE CAMBIO").Select
Application.Visible = True
        ActiveWindow.Zoom = 100
        ActiveWindow.SmallScroll Down:=-15
             ExecuteExcel4Macro ("show.toolbar(""ribbon"",0)")
        ActiveWindow.SmallScroll Down:=-15
    ActiveWindow.DisplayHorizontalScrollBar = False
        
 Me.Hide
End Sub

Private Sub CommandButton27_Click()
SOLICITUDTC.Show
End Sub

Private Sub CommandButton28_Click()
SOLICITUDCP.Show
End Sub

Private Sub CommandButton3_Click()
CANC.Show

End Sub

Private Sub CommandButton30_Click()
SEGUROS.Show
End Sub

Private Sub CommandButton31_Click()
RECLAMOS.Show
End Sub

Private Sub CommandButton4_Click()
CHPA.Show

End Sub
Private Sub CommandButton5_Click()
PASE.Show

End Sub
Private Sub CommandButton7_Click()
PAGO.Show

End Sub

Private Sub Frame1_activate()
Application.ActiveWindow.WindowState = xlMaximized
End Sub
Private Sub Image3_click()
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = False
Sheets("ULTIMO REGISTRO").Visible = False
Sheets("TIPO DE CAMBIO").Visible = False
Sheets("ULTIMA CUENTA").Visible = False
Sheets("BASE CUENTAS").Visible = False
Sheets("BUSC TARJETA").Visible = False
SALIDA.Show
Me.Hide
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
With Sheets("ULTIMO REGISTRO")
Application.ScreenUpdating = False
ListBox1.List = .Range("a1", "E1").Value
Me.ListBox1.ColumnWidths = "90 pt; 0 pt;200 pt; 0 pt;50 pt"

End With



End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub ListBox1_Change()

End Sub

Private Sub ListBox1_Click()

End Sub



Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
With Sheets("ULTIMO REGISTRO")

ListBox1.List = .Range("a1", "E1").Value
Me.ListBox1.ColumnWidths = "90 pt; 0 pt;200 pt; 0 pt;50 pt"

End With

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
TextBox3.Text = UCase(TextBox3.Text)

If TextBox3.Text = "COME" Then
COME.Show
TextBox3.Text = ""
Else
If TextBox3.Text = "VEME" Then
VEME.Show
TextBox3.Text = ""
Else
If TextBox3.Text = "RETI" Then
RETI.Show
TextBox3.Text = ""
Else
If TextBox3.Text = "CANC" Then
CANC.Show
TextBox3.Text = ""
Else
If TextBox3.Text = "CHPA" Then
CHPA.Show
TextBox3.Text = ""
Else
If TextBox3.Text = "PASE" Then
PASE.Show
TextBox3.Text = ""
    Else
    If TextBox3.Text = "DEPO" Then
    DEPO.Show
    TextBox3.Text = ""
        Else
        If TextBox3.Text = "DIRE" Then
        DIRE.Show
        TextBox3.Text = ""
        Else
        If TextBox3.Text = "DIEN" Then
        DIEN.Show
        TextBox3.Text = ""
        Else
               If TextBox3.Text = "PAGO" Then
        PAGO.Show
        TextBox3.Text = ""
            Else
            If TextBox3.Text = "EMIS" Then
            EMIS.Show
            TextBox3.Text = ""
                Else
                If TextBox3.Text = "COBR" Then
                COBR.Show
                TextBox3.Text = ""
                    Else
                    If TextBox3.Text = "PICA" Then
                    PICA.Show
                    TextBox3.Text = ""
                        Else
                        If TextBox3.Text = "CCFI" Then
                        CCFI.Show
                        TextBox3.Text = ""
                        End If
                    End If
                End If
            End If
    End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Initialize()

Label9.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
bState = False
End Sub

Private Sub UserForm_Terminate()
SALIDA.Show
End Sub

Private Sub UserForm_activate()
Sheets("REPORTE MONETARIO").Visible = True
Sheets("CARACTERÍSTICAS OPERATIVAS").Visible = True
Sheets("ULTIMO REGISTRO").Visible = True
Sheets("TIPO DE CAMBIO").Visible = True
Sheets("ULTIMA CUENTA").Visible = True
Sheets("BASE CUENTAS").Visible = True

Sheets("INICIO").Visible = True
Sheets("REPORTE MONETARIO").Select
TextBox1.Text = Cells(4, 2)
TextBox2.Text = Cells(2, 2)
TextBox4.Text = Cells(3, 2)

If Cells(3, 5) = "VERDADERO" Then
    CommandButton22.Enabled = True
    CommandButton16.Enabled = True
    CommandButton17.Enabled = True
    CommandButton2.Enabled = True
    CommandButton5.Enabled = True
    CommandButton7.Enabled = True
    CommandButton11.Enabled = True
    CommandButton14.Enabled = True
    CommandButton1.Enabled = True
    CommandButton4.Enabled = True
    CommandButton3.Enabled = True
    CommandButton12.Enabled = True
    CommandButton15.Enabled = True
Else
    CommandButton22.Enabled = False
    CommandButton16.Enabled = False
    CommandButton17.Enabled = False
    CommandButton2.Enabled = False
    CommandButton5.Enabled = False
    CommandButton7.Enabled = False
    CommandButton11.Enabled = False
    CommandButton14.Enabled = False
    CommandButton1.Enabled = False
    CommandButton4.Enabled = False
    CommandButton3.Enabled = False
    CommandButton12.Enabled = False
    CommandButton15.Enabled = False

End If


Application.ScreenUpdating = False
Application.Visible = True
ActiveWindow.Zoom = 150
Label8.Caption = DateValue(Now) + Time
Application.Visible = False
Application.ScreenUpdating = False

If TextBox4.Text <> "00000000" Then
CommandButton29.Visible = False

With ActiveWindow
    .DisplayHorizontalScrollBar = False
        .DisplayGridlines = False
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
    End With
    ExecuteExcel4Macro ("show.toolbar(""ribbon"",0)")
End If





End Sub
