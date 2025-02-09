Attribute VB_Name = "Guardar"
Sub Guardar()
    
'Macro que evita que se repitan el doc. Identidad en la base de datos

    EliminarFila.eliminar_fila

    Set h1 = Sheets("Registro") 'Colocar el nombre de la hoja donde est� el dato que se quiere evaluar.
    Set h2 = Sheets("DATOS") 'Colocar el nombre de la hoja donde se encuentran los registros para ser comparado con el dato mencionado m�s arriba.
    '
    If h1.[H7] = "" Then 'Entre los corchetes [D13] colocar la celda donde est� el dato que se quiere evaluar.
        MsgBox "Falta colocar el documento de identidad", vbExclamation, "GUARDAR" 'Entre comillas mensaje que se muestra si no hay datos en la celda definida m�s arriba.
        Exit Sub
    End If
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
    Set b = h2.Columns("A").Find(h1.[H7], lookat:=xlWhole) 'Cambiar las "D" por la columna donde se encuentran sus registros a evaluar, y cambiar [D13] por la celda que se quiere evaluar en la hoja principal.
    If Not b Is Nothing Then
        MsgBox "El usuario con este documeto ya existe", vbCritical, "GUARDAR" 'Entre las primeras comillas mensaje que se muestra si el dato evaluado existe en los registros.
        Exit Sub
    End If
    
    'A continuaci�n colocar el nombre de las macros (registro).    

    Registrar.registro
    
    MsgBox "El dato se guard�", vbInformation, "GUARDAR" 'Entre las primeras comillas mensaje que se muestra si su macro se ejecut� satisfactoriamente.
    
    'A continuaci�n colocar el nombre de la macro limpiar    

    Limpiar.Limpiar
    
    'Fin de la macro
End Sub

