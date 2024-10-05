Attribute VB_Name = "Módulo12"
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    ActiveWindow.SmallScroll Down:=-30
    Range("C16").Select
    ActiveWindow.SmallScroll Down:=6
    Union(Range( _
        "C40:J41,C16,E16,C17:J17,C18:J18,C19:D19,F19,I19,C20,C21,F20,I20,C26,E26:F26,C27:J27,C28:J28,C29:D29,F29,I29,C30,F30,I30,C31:J32,C35,E35:F35,C36:J36,C37:J37,C38:D38,C39,F38,F39,I38" _
        ), Range("I39")).Select
    Range("C40").Activate
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    Union(Range( _
        "C40:J41,B43:J45,B47:J49,C16,E16,C17:J17,C18:J18,C19:D19,F19,I19,C20,C21,F20,I20,C26,E26:F26,C27:J27,C28:J28,C29:D29,F29,I29,C30,F30,I30,C31:J32,C35,E35:F35,C36:J36,C37:J37,C38:D38,C39,F38" _
        ), Range("F39,I38,I39")).Select
    Range("B47").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-27
    Range("C16").Select
End Sub
