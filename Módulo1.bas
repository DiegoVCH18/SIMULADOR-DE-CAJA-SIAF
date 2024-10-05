Attribute VB_Name = "Módulo1"


Sub inputbox_Password(El_Form As Form, Caracter As String)

m_ASC = Asc(Caracter)

Call SetTimer(El_Form.hwnd, &H5000&, 100, AddressOf TimerProc)

End Sub


Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, _
ByVal dwTime As Long)

Dim Handle_InputBox As Long

'Captura el handle del textBox del InputBox
Handle_InputBox = FindWindowEx(FindWindow("#32770", App.Title), 0, "Edit", "")

'Le establece el PasswordChar
Call SendMessageLongRef(Handle_InputBox, &HCC&, m_ASC, 0)
'Finaliza el Timer
Call KillTimer(hwnd, idEvent)

End Sub


