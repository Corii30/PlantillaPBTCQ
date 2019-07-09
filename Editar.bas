Attribute VB_Name = "Editar"
Sub buscar()
Attribute buscar.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Range("H5").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[2]C,DATOS!R6C1:R1048576C15,2,0),"""")"
    Range("H5").Select
    Selection.Copy
    Range("H9,H11,H13,H15,H17,K5,K9,K11,K13,K15,K17").Select
    Range("K17").Activate
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H9").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-2]C,DATOS!R6C1:R1048576C15,3,0),"""")"
    Range("H11").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-4]C,DATOS!R6C1:R1048576C15,4,0),"""")"
    Range("H13").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-6]C,DATOS!R6C1:R1048576C15,5,0),"""")"
    Range("H15").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-8]C,DATOS!R6C1:R1048576C15,6,0),"""")"
    Range("H17").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-10]C,DATOS!R6C1:R1048576C15,7,0),"""")"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],DATOS!R6C1:R1048576C15,9,0),"""")"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-2]C[-3],DATOS!R6C1:R1048576C15,10,0),"""")"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-4]C[-3],DATOS!R6C1:R1048576C15,14,0),"""")"
    Range("K13").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-6]C[-3],DATOS!R6C1:R1048576C15,11,0),"""")"
    Range("K15").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-8]C[-3],DATOS!R6C1:R1048576C15,12,0),"""")"
    Range("K17").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-10]C[-3],DATOS!R6C1:R1048576C15,13,0),"""")"
    Range("H5").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H9").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H13").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H15").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H17").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K17").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H7").Select
    MinusMayus.convminus
    Application.ScreenUpdating = True
End Sub



