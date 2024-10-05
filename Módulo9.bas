Attribute VB_Name = "Módulo9"
Sub tarjetas()
Attribute tarjetas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' tarjetas Macro
'

'
    Range("T13:T17").Select
    Selection.Copy
    Sheets("TIPO DE CAMBIO").Select
    ActiveWindow.SmallScroll Down:=432
    Range("B450").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    ActiveWindow.SmallScroll Down:=-504
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("B2:B450"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Sheets("SOLICITUD TC").Select
End Sub
