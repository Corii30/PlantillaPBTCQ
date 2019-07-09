Attribute VB_Name = "EliminarFila"
Sub eliminar_fila()
Attribute eliminar_fila.VB_ProcData.VB_Invoke_Func = " \n14"
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


