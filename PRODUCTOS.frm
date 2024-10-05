VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PRODUCTOS 
   Caption         =   "Productos por cliente"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   OleObjectBlob   =   "PRODUCTOS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "PRODUCTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, items, xEmpleado

Private Sub btn_Buscar_Click()

        If Me.TXT_BUSCAR.Value = Empty Then
            MsgBox "Escriba un registro para buscar", vbExclamation, "SIAF"
            Me.ListBox1.Clear
            Me.TXT_BUSCAR.SetFocus
            Exit Sub
        End If

items = Range("Tabla2").CurrentRegion.Rows.Count
        For i = 2 To items
            If LCase(Cells(i, 2).Value) Like "*" & LCase(Me.TXT_BUSCAR.Value) & "*" Then
                Me.ListBox1.AddItem Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Cells(i, 2)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Cells(i, 6)
            End If
        Next i
        Me.TXT_BUSCAR.SetFocus
        Me.TXT_BUSCAR.SelStart = 0
        Me.TXT_BUSCAR.SelLength = Len(Me.TXT_BUSCAR.Text)
Exit Sub
TXT_BUSCAR.Text = ""
End Sub

Private Sub btn_Eliminar_Click()

CANC.TextBox12.Text = TextBox2
CANC.TextBox7.Text = TextBox3
CANC.ComboBox1.Text = TextBox4
DEPO.TextBox12.Text = TextBox2
DEPO.TextBox7.Text = TextBox3
DEPO.ComboBox1.Text = TextBox4
RETI.TextBox12.Text = TextBox2
RETI.TextBox7.Text = TextBox3
RETI.ComboBox1.Text = TextBox4
PAGO.TextBox7.Text = TextBox2

RECLAMOS.TextBox9.Text = TextBox2
RECLAMOS.TextBox7.Text = TextBox3
Sheets("ULTIMA CUENTA").Select
Cells(1, 13) = TextBox3.Text
CANC.ComboBox4.Text = Cells(1, 14)
DEPO.ComboBox4.Text = Cells(1, 14)
RETI.ComboBox4.Text = Cells(1, 14)

Unload Me

End Sub

Private Sub Label19_Click()

End Sub

Private Sub ListBox1_Click()
Dim i As Long
Dim dato As Integer
For i = 0 To ListBox1.ListCount - 0
If ListBox1.Selected(i) Then dato = ListBox1.List(i)
Next i
Sheets("TIPO DE CAMBIO").Cells(dato + 1, 1).Activate
TextBox1 = ActiveCell.Offset(0, 1)
TextBox4 = ActiveCell.Offset(0, 3)
TextBox2 = ActiveCell.Offset(0, 4)
TextBox3 = ActiveCell.Offset(0, 5)
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub txt_Buscar_Change()

End Sub

Private Sub UserForm_activate()
Application.Visible = True
Sheets("TIPO DE CAMBIO").Select
fin = Application.CountA(Sheets("TIPO DE CAMBIO").Range("B:B"))
Sheets("TIPO DE CAMBIO").Range("A2:A" & fin).Select
Selection.Clear
ReDim Mirango(1 To fin)
    For i = 1 To fin:
        Mirango(i) = i
Next i
Worksheets("TIPO DE CAMBIO").Range("A2:A" & UBound(Mirango)).Value = _
Application.WorksheetFunction.Transpose(Mirango)
Worksheets("TIPO DE CAMBIO").Range("A1").Select

'Le digo cuántas columnas
    ListBox1.ColumnCount = 6
    
    'Asigno el ancho a cada columna
    Me.ListBox1.ColumnWidths = "40 PT;50 pt;140 pt;100 pt;130 pt;120 PT"
        'El origen de los datos es la Tabla1
         '   ListBox1.RowSource = "Tabla1"
End Sub



