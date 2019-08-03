Sub EasyStock()

Dim ticker As String
Dim total_stock As Double
total_stock = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
      ticker = Cells(i, 1).Value

      total_stock = total_stock + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = ticker

      Range("J" & Summary_Table_Row).Value = total_stock

      Summary_Table_Row = Summary_Table_Row + 1
      
      total_stock = 0

    Else

      total_stock = total_stock + Cells(i, 7).Value

    End If
 
Next i

End Sub