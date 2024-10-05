Attribute VB_Name = "Módulo11"
Sub pruebahorro()
Attribute pruebahorro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' pruebahorro Macro
'

'
    Range("Q13:Q17").Select
    Selection.Copy
    Sheets("TIPO DE CAMBIO").Select
    ActiveWindow.SmallScroll Down:=432
    Range("B450").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    ActiveWindow.SmallScroll Down:=-488
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").ListObjects("Tabla2").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").ListObjects("Tabla2").Sort. _
        SortFields.Add2 Key:=Range("Tabla2[[#All],[DNI]]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TIPO DE CAMBIO").ListObjects("Tabla2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=0
    Range("A3").Select
    ActiveWindow.SmallScroll Down:=8
    Range("A34").Select
    Selection.AutoFill Destination:=Range("A34:A450")
    Range("A34:A450").Select
    Range("A34").Select
    Selection.AutoFill Destination:=Range("A20:A34"), Type:=xlFillDefault
    Range("A20:A34").Select
    ActiveWindow.SmallScroll Down:=8
    Sheets("CARTILLA CUENTA").Select
    Range("Q13").Select
End Sub
