Sub ExtractSheetName_Multiple()
   Dim path As String
   Dim wb As Workbook
   Dim ws As Worksheet
   Dim sel As Range

   'Obtener la selección
   Set sel = Selection

   'Iterar a través de las celdas seleccionadas
   For Each cell In sel
      path = cell.Value

      'Abrir el archivo de Excel
      Set wb = Workbooks.Open(path)

      'Obtener la primera hoja del archivo
      Set ws = wb.Sheets(1)

      'Insertar el nombre de la hoja en la siguiente columna
      cell.Offset(0, 1).Value = ws.Name

      'Cerrar el archivo
      wb.Close False
   Next
End Sub
