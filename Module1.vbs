Attribute VB_Name = "Module1"
Sub Columns()
'Changes to all Worksheets

For Each ws In Worksheets


' Create 4 new columns
ws.Range("I1:L1").EntireColumn.Insert
 
' Name each column

Dim name: name = Split("Ticker,Yearly Change,Percent Change,Total Stock Volume", ",")

ws.Range("I1").Resize(1, UBound(name) + 1) = name
 
 
Next ws

 
 

 

End Sub

