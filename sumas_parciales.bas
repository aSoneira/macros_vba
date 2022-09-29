Sub sumas()
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
