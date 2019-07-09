Attribute VB_Name = "Limpiar"
Sub Limpiar()
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Range("K17").Select
    Selection.ClearContents
    Range("K15").Select
    Selection.ClearContents
    Range("K13").Select
    Selection.ClearContents
    Range("K11").Select
    Selection.ClearContents
    Range("K9").Select
    Selection.ClearContents
    Range("K7").Select
    Selection.ClearContents
    Range("H17").Select
    Selection.ClearContents
    Range("H15").Select
    Selection.ClearContents
    Range("H13").Select
    Selection.ClearContents
    Range("H11").Select
    Selection.ClearContents
    Range("H9").Select
    Selection.ClearContents
    Range("H7").Select
    Selection.ClearContents
    
    AumentoCod.aumento_codigo
    Fecha.Fecha
    
    Application.ScreenUpdating = True
End Sub
