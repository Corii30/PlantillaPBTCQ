Attribute VB_Name = "Eliminar_fila"
Sub Eliminar_fila()
    Attribute aumento_codigo.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Sheets("DATOS").Select
    Rows("1048576:1048576").Select
    Selection.Delete Shift:=xlUp
    Selection.End(xlUp).Select
    Range("A7").Select
    Sheets("Registro").Select
    Range("H7").Select
    Application.ScreenUpdating = True
End Sub


