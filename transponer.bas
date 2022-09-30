Sub transponer()
  fila = 20
  col = 10
  largo = 12
  For i = 0 To largo
      i_t1 = fila
      i_t2 = fila + 21
      j_t1 = 3 + i
      j_t2 = 3 + i
      
      prev = "R" & i_t1 & "C" & j_t1
      seg = "R" & i_t2 & "C" & j_t2
      Cells(60 + i, col).FormulaR1C1 = "=" & seg & "/" & prev
        
  Next i
End Sub
