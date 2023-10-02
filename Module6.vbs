Attribute VB_Name = "Module6"
Sub Solve_for_Greatest()
    Dim ws As Worksheet
    Dim r As Range
    Dim m As Double
    Dim lastRow As Long
    Dim ticker As String
     ticker = 9
    For Each ws In Worksheets
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        
        Set r = ws.Range("K2:K" & lastRow)
        
        ' Find the Greatest increase %
        m = Application.WorksheetFunction.Max(r)
        
        ' Place the maximum value in cell R2C18 (row 2, column 18)
        ws.Cells(2, 18).Value = m * 100
        
        'Find the ticker associated with the Max
        
        
        Set r = ws.Range("K2:K" & lastRow)
        
        ' Find the Greatest Decrease %
        m = Application.WorksheetFunction.Min(r)
        
        ' Place the maximum value in cell R2C18 (row 2, column 18)
        ws.Cells(3, 18).Value = m * 100
        
        'Find the ticker associated with the Min
        
        
    Set r = ws.Range("L2:L" & lastRow)
        
        ' Find the Greatest Total Volume
        m = Application.WorksheetFunction.Max(r)
        
        ' Place the maximum value in cell R2C18 (row 2, column 18)
        ws.Cells(4, 18).Value = m * 100
        
        'Find the ticker associated with the Max
   
    
    Next ws
End Sub

