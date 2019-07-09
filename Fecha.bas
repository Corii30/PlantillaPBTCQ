Attribute VB_Name = "Fecha"
Sub Fecha()
    Application.ScreenUpdating = False
    Sheets("Registro").Select Range("K11").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Sheets("Registro").Select Range("K11").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H7").Select
    Application.ScreenUpdating = True
End Sub

