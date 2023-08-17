Attribute VB_Name = "Module1"
Sub Script():

Dim ws As Worksheet
'Loop through all sheets in ws

For Each ws In ThisWorkbook.Sheets
ws.Activate

'Last Row
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Dim all variables
Dim ticker As String
Dim open_col As Double
Dim close_col As Double
Dim yearly_change As Double
Dim firstRow As Double
Dim precentage_change As Long
Dim total_stock As Double
Dim max_ticker As String
Dim min_ticker As String
Dim volume_ticker As String
Dim max_value As Double
Dim min_value As Double
Dim max_volume As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
firstRow = 2
total_stock = 0
        max_value = 0
        min_value = 0
        max_volume = 0
        
'Loop through all rows
For i = 2 To lastRow

'Set cycle to find different categoriesof Tickers
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set the ticker value
ticker = Cells(i, 1).Value

'Set header for ticker
 ws.Cells(1, 9).Value = "Ticker"
 
 'Put the values of each ticker into summary
 Range("I" & Summary_Table_Row).Value = ticker
 
'Header for yearly change
 ws.Cells(1, 10).Value = "Yearly Change"
 
'name header total stock value
 ws.Cells(1, 12).Value = "Total stock values"
 
'Yearly Change
close_col = Cells(i, 6).Value
open_col = Cells(firstRow, 3).Value
yearly_change = close_col - open_col
  
'Print the yearly change Summary Table
 Range("J" & Summary_Table_Row).Value = yearly_change
 
'Header for precentage change
 ws.Cells(1, 11).Value = "Precentage Change"
 
'calculate percentage change
percentage_change = (yearly_change / open_col)

If percentage_change > max_value Then
max_value = percentage_change
max_ticker = ticker
End If

If percentage_change < min_value Then
min_value = percentage_change
min_ticker = ticker
End If

'Total stock values
total_stock = total_stock + Cells(i, 7).Value

If total_stock > max_volume Then
max_volume = total_stock
volume_ticker = ticker

End If

Cells(2, 16).Value = max_ticker
Cells(2, 17).Value = max_value
Cells(3, 16).Value = min_ticker
Cells(3, 17).Value = min_value
Cells(4, 16).Value = volume_ticker
Cells(4, 17).Value = max_volume
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"

'Print the percentage change in the Summary Table Column K
Range("K" & Summary_Table_Row).Value = percentage_change
Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

'Total stockvalue
Range("L" & Summary_Table_Row).Value = total_stock

'Conditional formatting to match index colour for yearly change
If yearly_change > 0 Then
Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

Else
Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

End If

'Conditional formating to match index colour for precentage change
If percentage_change > 0 Then
Range("K" & Summary_Table_Row).Interior.ColorIndex = 4

Else
Range("K" & Summary_Table_Row).Interior.ColorIndex = 3

End If

Cells(1, 15).Value = "Metric"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


total_stock = 0
firstRow = i + 1

Summary_Table_Row = Summary_Table_Row + 1

Else

total_stock = total_stock + Cells(i, 7).Value
End If
Next i
Next ws
End Sub

