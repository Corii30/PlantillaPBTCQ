Attribute VB_Name = "Registrar"
Sub registro()
    Application.ScreenUpdating = False
    Sheets("DATOS").Select
    Rows("7:7").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Sheets("Registro").Select
    Selection.Copy
    Sheets("DATOS").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Registro").Select
    Application.CutCopyMode = False
    Range("H7").Select
    Selection.Copy
    Sheets("DATOS").Select
    Range("A7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Sheets("Registro").Select
    Application.CutCopyMode = False
    Range("H5,H9,H11,H13,H15,H17").Select
    Range("H17").Activate
    Selection.Copy
    Sheets("DATOS").Select
    Range("B7:G7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Sheets("Registro").Select
    Application.CutCopyMode = False
    Range("K5,K7,K9").Select
    Selection.Copy
    Sheets("DATOS").Select
    Range("H7:J7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Sheets("Registro").Select
    Application.CutCopyMode = False
    Range("K13,K15,K17").Select
    Range("K17").Activate
    Selection.Copy
    Sheets("DATOS").Select
    Range("K7:M7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Sheets("Registro").Select
    Application.CutCopyMode = False
    Range("K11").Select
    Selection.Copy
    Sheets("DATOS").Select
    Range("N7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A7:O7").Select
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    MinusMayus.convmays
    Range("A7").Select
    Sheets("Registro").Select
    Range("H7").Select
    Application.ScreenUpdating = True
End Sub

