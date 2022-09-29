Attribute VB_Name = "Subtotales"
Sub Sumatorios()

'Ojo con la aplicación, que se interrumpe el bucle en caso de filas vacías

'Además, tener en cuenta que todavía hay que definir la columna de marcas
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
columnaMarcas = 1
marcaObra = "o"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Application.ScreenUpdating = False

Dim calc_mode As Variant
    calc_mode = Application.Calculation
    Application.Calculation = xlCalculationManual

'*************************************************************
'Búsqueda de filaInicio y filaFin

For Each selec In Selection
    If IsEmpty(filaInicio) Then
        filaInicio = selec.Row
    End If
    If IsEmpty(filaFin) Then
        filaFin = selec.Row
    End If
'    If IsEmpty(columnaInicio) Then
'        columnaInicio = selec.Column
'    End If
'    If IsEmpty(columnaFin) Then
'        columnaFin = selec.Column
'    End If
   
    If selec.Row < filaInicio Then
        filaInicio = selec.Row
    End If
    If selec.Row > filaFin Then
        filaFin = selec.Row
    End If
    
'    If selec.Column < columnaInicio Then
'        columnaInicio = selec.Column
'    End If
'    If selec.Row > columna.Fin Then
'        columnaFin = selec.Column
'    End If

Next

'-------------------------------------------------------------
Dim a As Integer        'Coordenada fila de la celda actual
Dim b As Integer        'Coordenada columna de la celda actual
Dim formula As String   'fórmula a insertar

For Each sel In Selection

a = sel.Row
b = sel.Column

If IsNumeric(Cells(a, columnaMarcas).Value) Then            'la celda seleccionada es un capítulo
    If IsNumeric(Cells(a + 1, columnaMarcas).Value) Then    'la fila de debajo también es un capítulo
        formula = "="
        For i = a + 1 To filaFin
            If IsNumeric(Cells(i, columnaMarcas).Value) Then            'la fila detectada es un capítulo
                If Cells(i, columnaMarcas).Value = Cells(a, columnaMarcas).Value + 1 Then   'y dicha fila es justo de un nivel inferior
                    formula = formula & "+R[" & (i - a) & "]C[]"
                ElseIf Cells(i, columnaMarcas).Value <= Cells(a, columnaMarcas).Value Then  ' y dicha fila es de un nivel igual o superior
                    Exit For
                End If
           
            ElseIf Cells(i, columnaMarcas).Value = "" Then      'si la fila detectada no encuentra marca alguna
                Exit For
            End If
            
            
        Next i
        
        '!!!!!!!!!!!!!!!!ESCRITURA de la fórmula en celda
        If (Not formula = "=") Then
        sel.FormulaR1C1 = formula
        End If
        '!!!!!!!!!!!!!!!!
        
    ElseIf (Cells(a + 1, columnaMarcas).Value = marcaObra) Then   'la fila de debajo es una unidad de obra
        formula = "=sum(R[1]C[]:R["
        For i = a + 1 To filaFin
            If Not Cells(i, columnaMarcas).Value = marcaObra Then    'si la fila detectada NO es una ud de obra
                formula = formula & (i - 1 - a) & "]C[])"
                Exit For
            End If
            
            If i = filaFin Then         'si el bucle llega al final de la selección
                formula = formula & (i - a) & "]C[])"
                Exit For
            End If
            
        Next i
        sel.FormulaR1C1 = formula     'ESCRITURA de la fórmula en celda
        
    End If

ElseIf (Cells(a, columnaMarcas).Value = marcaObra) Then
    GoTo finRutina
End If
finRutina:

Next

'*************************************************************
Application.ScreenUpdating = True

Application.Calculate
Application.Calculation = calc_mode

End Sub
