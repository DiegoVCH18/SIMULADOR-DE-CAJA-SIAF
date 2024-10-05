Attribute VB_Name = "Módulo6"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveWindow.DisplayHeadings = True
    Cells.Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    Range("A9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    Range("F21").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    ActiveSheet.EnableSelection = xlUnlockedCells
    Range("A13").Select
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    ActiveWindow.DisplayHeadings = True
    Cells.Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    Range("A9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    Range("B15").Select
    ActiveWindow.SmallScroll Down:=-15
    ActiveWindow.DisplayHeadings = False
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Range("A14").Select
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    ActiveSheet.Unprotect
    Range("A9").Select
End Sub
