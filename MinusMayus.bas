Attribute VB_Name = "MinusMayus"
Sub convmays()

    Application.ScreenUpdating = False
    Set rgColA = Union(Sheets("DATOS").Range("c7:e7"), Sheets("DATOS").Range("h7:j7"))
    Dim rg As Range
    For Each rg In rgColA.Cells
        rg.Value = UCase(rg.Text)
    Next
    Set rg2 = Sheets("DATOS").Range("k7")
    If Not rg2.Value = "" Then
        rg2.Value = UCase(rg2.Text)
    Else
        rg2.Value = ""
    End If

    Application.ScreenUpdating = True
End Sub

Sub convminus()

    Application.ScreenUpdating = False
    Set rgColA = Union(Sheets("Registro").Range("h9:h15"), Sheets("Registro").Range("k7:k11"))
    Dim rg As Range
    For Each rg In rgColA.Cells
        rg.Value = LCase(rg.Text)
    Next
    Set rg2 = Sheets("Registro").Range("k15")
    If Not rg2.Value = "" Then
        rg2.Value = LCase(rg2.Text)
    Else
        rg2.Value = ""
    End If

    Application.ScreenUpdating = True

End Sub


