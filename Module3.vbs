Attribute VB_Name = "Module3"
 Sub Yearly_change()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker_name As String
    Dim year_open As Double
    Dim year_close As Double
    Dim year_change As Double
  
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim consolidated_info_row As Integer
        consolidated_info_row = 2
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set the ticker name
                ticker_name = ws.Cells(i, 1).Value
                
                ' Capture start value
                year_open = ws.Cells(i - 250, 3).Value
                year_close = ws.Cells(i, 6).Value
                
                ' Calculate Yearly change of values
                year_change = year_close - year_open ' Calculate the difference
                
                ' Print the Yearly Change in the J column
                ws.Cells(consolidated_info_row, 10).Value = year_change
                
                ' move to the next row
                consolidated_info_row = consolidated_info_row + 1
                
                ' Reset the Data
                year_open = 0
                year_close = 0
                year_change = 0
            
              
End If
        ' Change color of row based on year_change value
        If ws.Cells(i, 10) <> "" Then
            
            If ws.Cells(i, 10) < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 4

    End If
    
    
End If

Next i
    
    
    
Next ws


End Sub
