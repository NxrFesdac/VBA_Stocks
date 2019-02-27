Sub Tarea()

Dim L As Long
Dim LastRow As Long
Dim WS_Count As Integer
Dim I As Integer
Dim vol As Double
Dim tabla As Long
Dim abierto As Double
Dim precio As Double
Dim mayor As Double
Dim nombre As String


abierto = 2
tabla = 2
WS_Count = ActiveWorkbook.Worksheets.Count


For I = 1 To WS_Count

LastRow = Worksheets(I).Cells(Worksheets(I).Rows.Count, 1).End(xlUp).Row


For L = 2 To LastRow
  If Worksheets(I).Cells(L + 1, 1).Value <> Worksheets(I).Cells(L, 1).Value Then
      Worksheets(I).Cells(tabla, 9).Value = Worksheets(I).Cells(L, 1).Value
      vol = vol + Worksheets(I).Cells(L, 7).Value
      Worksheets(I).Cells(tabla, 10).Value = vol
      vol = 0
      If Worksheets(I).Cells(abierto, 3).Value - Worksheets(I).Cells(L, 6).Value > 0 Then
         Worksheets(I).Cells(tabla, 12).Interior.ColorIndex = 4
      Else
         Worksheets(I).Cells(tabla, 12).Interior.ColorIndex = 3
      End If
      Worksheets(I).Cells(tabla, 12).Value = Worksheets(I).Cells(abierto, 3).Value - Worksheets(I).Cells(L, 6).Value
      If Worksheets(I).Cells(abierto, 3).Value = 0 Then
         Worksheets(I).Cells(tabla, 11).Value = Worksheets(I).Cells(L, 6).Value * -1
      Else
         Worksheets(I).Cells(tabla, 11).Value = ((Worksheets(I).Cells(L, 6).Value - Worksheets(I).Cells(abierto, 3).Value) / Worksheets(I).Cells(abierto, 3).Value)
      End If
      Worksheets(I).Cells(tabla, 11).Style = "Percent"
      tabla = tabla + 1
      abierto = (L + 1)

  Else
      vol = vol + Worksheets(I).Cells(L, 7).Value
  End If
Next L

tabla = 2
vol = 0
abierto = 2

LastRow = Worksheets(I).Cells(Worksheets(I).Rows.Count, 9).End(xlUp).Row

For L = 2 To LastRow
   If Worksheets(I).Cells(L, 10) >= mayor Then
       mayor = Worksheets(I).Cells(L, 10).Value
       nombre = Worksheets(I).Cells(L, 9).Value
   End If
Next L

Worksheets(I).Cells(2, 15).Value = nombre
Worksheets(I).Cells(2, 16).Value = mayor


nombre = ""
mayor = 0

For L = 2 To LastRow
   If Worksheets(I).Cells(L, 11) >= mayor Then
       mayor = Worksheets(I).Cells(L, 11).Value
       nombre = Worksheets(I).Cells(L, 9).Value
   End If
Next L

Worksheets(I).Cells(3, 15).Value = nombre
Worksheets(I).Cells(3, 16).Value = mayor

nombre = ""
mayor = 0

For L = 2 To LastRow
   If Worksheets(I).Cells(L, 11) <= mayor Then
       mayor = Worksheets(I).Cells(L, 11).Value
       nombre = Worksheets(I).Cells(L, 9).Value
   End If
Next L

Worksheets(I).Cells(4, 15).Value = nombre
Worksheets(I).Cells(4, 16).Value = mayor

nombre = ""
mayor = 0

Worksheets(I).Cells(2, 16).NumberFormat = "General"
Worksheets(I).Cells(3, 16).Style = "Percent"
Worksheets(I).Cells(4, 16).Style = "Percent"

Worksheets(I).Cells(1, 15).Value = "Ticket"
Worksheets(I).Cells(1, 16).Value = "Value"
Worksheets(I).Cells(2, 14).Value = "Greatest Volume"
Worksheets(I).Cells(3, 14).Value = "Greatest % Increase"
Worksheets(I).Cells(4, 14).Value = "Greatest % Decrease"
Worksheets(I).Cells(1, 9).Value = "Ticket"
Worksheets(I).Cells(1, 10).Value = "Volume"
Worksheets(I).Cells(1, 11).Value = "Percentage change"
Worksheets(I).Cells(1, 12).Value = "Price Change"



Next I


End Sub
