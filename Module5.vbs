Attribute VB_Name = "Module5"
Sub Create_Greatest_Columns()

For Each ws In Worksheets
' Insert new columns
ws.Range("Q1:R1").EntireColumn.Insert

' Name new columns
Dim name: name = Split("Ticker,Value", ",")

ws.Range("Q1").Resize(1, UBound(name) + 1) = name

' Insert categories in Rows

ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"

Next ws

End Sub

