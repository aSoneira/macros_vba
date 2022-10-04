'Ejecutar bucle sobre un rango de hojas seleccionadas
'Fuente: https://exceloffthegrid.com/loop-through-selected-sheets-with-vba/

For Each ws In ActiveWindow.SelectedSheets

    'Perform action.  E.g. hide selected worksheets
    ws.Visible = xlSheetVeryHidden

Next ws

