Attribute VB_Name = "Reemplazar"
Sub Reemplazar()

Set h1 = Sheets("Registro") 'Cambiar "principal" por la hoja que contenga el n� de registro que se desee reemplazar
Set h2 = Sheets("DATOS") 'Cambiar "registros" por la hoja que contenga la base de datos donde se encuentre el registro que se desea reemplazar

cf = MsgBox("Desea reemplazar el registro?", vbInformation + vbYesNo, "AVISO") 'Mensaje
If cf = vbYes Then

'Estas cuatros l�neas seguidas son para que la macro se detenga si falta informaci�n en la celda especificada
If h1.[H7] = "" Then 'Cambiar [D11] por la celda que contenga el n� de registro que se quiere reemplazar
MsgBox "Ingresar el documento de identidad" 'Mensaje que se muestra si la celda B2 est� vac�a
Exit Sub
End If

If h1.[H9] = "" Then 'Cambiar [D12] por la siguiente celda que se desee reemplazar
MsgBox "Falta el nombre del usario" 'Mensaje que se muestra si la celda D12 est� vac�a
Exit Sub
End If
'
If h1.[H15] = "" Then 'Cambiar [D13] por la siguiente celda que se desee reemplazar
MsgBox "Falta un n�mero de contacto" 'Mensaje que se muestra si la celda D13 est� vac�a
Exit Sub
End If
'
If h1.[K5] = "" Then 'Cambiar [D14] por la siguiente celda que se desee reemplazar
MsgBox "Falta la edad del usario" 'Mensaje que se muestra si la celda D14 est� vac�a
Exit Sub
End If

Set r = h2.Columns("A") 'Cambiar "B" por la columna donde se encuentra el n� de registro a reemplazar en la base de datos
Set b = r.Find(h1.[H7], lookat:=xlWhole) 'Cambiar [D11] por la celda donde esta el n� de registro que se desea reemplazar


If Not b Is Nothing Then 'Cambiar B por la columna donde est�n los datos que se desean reemplazar
    
    h2.Cells(b.Row, "C") = h1.[H9] 'Cambiar C por la columna donde est�n los datos que se desea reemplazar con los datos de la hoja "Principal". Cambiar [D12] por la celda seguida del n� de registro en la hoja "Principal" que se desea reemplazar
    h2.Cells(b.Row, "D") = h1.[H11]
    h2.Cells(b.Row, "E") = h1.[H13]
    h2.Cells(b.Row, "F") = h1.[H15]
    h2.Cells(b.Row, "G") = h1.[K5]
    h2.Cells(b.Row, "I") = h1.[K9]
    h2.Cells(b.Row, "J") = h1.[K11]
    h2.Cells(b.Row, "K") = h1.[K15]
    
    h2.Cells(b.Row, "M") = "=TODAY()"
    
    If h2.Cells(b.Row, "M").HasFormula Then
        h2.Cells(b.Row, "M").Value = h2.Cells(b.Row, "M").Value
    End If
    
    Set rgColA = Union(Range(h2.Cells(b.Row, 3), h2.Cells(b.Row, 5)), Range(h2.Cells(b.Row, 8), h2.Cells(b.Row, 10)))

    Dim rg As Range
    For Each rg In rgColA.Cells
        rg.Value = UCase(rg.Text)
    Next
    Set rg2 = h2.Cells(b.Row, 11)
    If Not rg2.Value = "" Then
        rg2.Value = UCase(rg2.Text)
    Else
        rg2.Value = ""
    End If
        
    MsgBox "Se ha reemplazado con �xito", vbInformation 'Mensaje
Else
    MsgBox "El doc.Identidad no existe", vbInformation 'Mensaje
    Exit Sub
End If
'A continuaci�n colocar la macro limpiar

Limpiar.Limpiar

End If
End Sub

