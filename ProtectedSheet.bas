Attribute VB_Name = "ProtectedSheet"
Sub protected_sheet()
Attribute protected_sheet.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("H7").Select
End Sub

Sub unprotected_sheet()
Attribute unprotected_sheet.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.Unprotect
    Range("H7").Select
End Sub


