Sub creaCarpetas()
  
  'Crear carpetas con la fecha de cada día del mes, siempre que este sea un día laborable
  Uri = "\\192.168.1.10\ejemplo"
  diames = 7
  mes = Format(diames, "00")
  For i = 1 To 31
    dia = Format(i, "00")
    If (Weekday((dia & "/" & mes & "/2022"), vbMonday) <= 5) Then
      
      ''Función principal, para la creación de carpetas:
      MkDir Uri & "\2022" & mes & dia
    
    End If
  Next i
  
  
End Sub
