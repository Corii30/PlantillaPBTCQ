Attribute VB_Name = "AumentoCod"
Sub aumento_codigo()
Attribute aumento_codigo.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Sheets("Registro").Select Range("H5").Select
    ActiveCell.FormulaR1C1 = "=COUNT(DATOS!R6C2:R1048576C2)+1"
    Sheets("Registro").Select Range("H5").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H7").Select
    Application.ScreenUpdating = True
End Sub


