Attribute VB_Name = "Editar"
Sub buscar()
Attribute buscar.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Range("H5").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[2]C,DATOS!R6C1:R1048576C12,2,0),"""")"
    Range("H5").Select
    Selection.Copy
    Range("H9,H11,H13,H15,K5,K9,K11,K13,K15").Select
    Range("K15").Activate
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("H9").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-2]C,DATOS!R6C1:R1048576C12,3,0),"""")"
    Range("H11").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-4]C,DATOS!R6C1:R1048576C12,4,0),"""")"
    Range("H13").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[2]C,DATOS!R6C1:R1048576C12,5,0),"""")"
    Range("H15").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-8]C,DATOS!R6C1:R1048576C12,6,0),"""")"
    Range("H13").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-6]C,DATOS!R6C1:R1048576C12,5,0),"""")"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[2]C[-3],DATOS!R6C1:R1048576C12,7,0),"""")"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-2]C[-3],DATOS!R6C1:R1048576C12,9,0),"""")"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-4]C[-3],DATOS!R6C1:R1048576C12,10,0),"""")"
    Range("K13").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-6]C[-3],DATOS!R6C1:R1048576C12,12,0),"""")"
    Range("K15").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(R[-8]C[-3],DATOS!R6C1:R1048576C12,11,0),"""")"
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
    Range("K5").Select
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
    Application.CutCopyMode = False
    Range("H7").Select
    
    convminus
    
    Application.ScreenUpdating = True
End Sub



