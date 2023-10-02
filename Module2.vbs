Attribute VB_Name = "Module2"

Sub consolidated_ticker()
' All worksheets

For Each ws In ThisWorkbook.Worksheets


' Define ticker and total stock volume

Dim ticker_name As String

Dim total_volume As Double
    total_volume = 0
    
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
' Define location of consolidated info

Dim consolidated_info_row As Integer
    consolidated_info_row = 2
    
    
        'Define loop target
 For i = 2 To lastRow
 
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      ticker_name = ws.Cells(i, 1).Value

      ' Add to the Stock Total
      total_volume = total_volume + ws.Cells(i, 7).Value

      ' Print the ticker name in the I column
      ws.Cells(consolidated_info_row, 9).Value = ticker_name

      ' Print the stock total in the L column
      ws.Cells(consolidated_info_row, 12).Value = total_volume

      ' move to next row
      consolidated_info_row = consolidated_info_row + 1
      
      ' Reset the Total Volume
      total_volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the tickers total stock volume
      total_volume = total_volume + ws.Cells(i, 7).Value
    

    

End If

Next i

Next ws

End Sub


