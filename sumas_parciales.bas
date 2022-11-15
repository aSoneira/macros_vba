Sub sumas_parciales()
For Each sel In Selection
  fila = 0
  col = -1
  paso = 1
  condicion = True
  Do While condicion = True
    fila = fila + paso
    If (Cells(sel.Row + fila, sel.Column + col).Value = "") Then
    condicion = False
      fila = fila - paso
      Exit Do
    End If
  Loop
  sel.FormulaR1C1 = "=sum(R[" & paso & "]C[" & col & "]:R[" & fila & "]C[" & col & "])"
Next
End Sub

'' ----------------------------------------------------------

'ALTERNATIVA
'Subrutina más simple, que permite detectar unidades de obra ("o") y 
'generar subtotales de las filas inferiores que cumplan una determinada condición 
' (en este caso, que la fila de la celda en cuestión empiece por "m")

Sub sumas_parciales_2()

For i = 6 To 6319
  If Cells(i, 1).Value = "o" Then
    inicio = i
    n = 1
    
    Do While Left(Cells(i + n, 1).Value, 1) = "m"
      n = n + 1
    Loop
    
    Cells(inicio, 9).FormulaR1C1 = "=sum(R[1]C:R[" & n - 1 & "]C)"
    i = i + n - 1
  End If
Next i

End Sub

