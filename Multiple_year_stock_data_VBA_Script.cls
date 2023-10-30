VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks_Analysis()

'Loop through all sheets
For Each ws In Worksheets

Dim Ticker_symbol As String
Dim Yearly_change As Double
Dim Percent_change As Double
Dim Total_stock_volume As Double
Total_stock_volume = 0
Dim Summary_table_row As Long
Summary_table_row = 2
Dim Opening_price As Double
Opening_price = Cells(2, 3).Value
Dim Closing_price As Double

'Counts the number of the row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all ticker symbol
For i = 2 To lastrow

'Check if we are still the same ticker symbol, if it is not
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Set the ticker symbol
    Ticker_symbol = ws.Cells(i, 1).Value
    
    'Calculate yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
    Closing_price = ws.Cells(i, 6).Value
    Yearly_change = Closing_price - Opening_price
    'Calculate the percentage change
    Percent_change = (Yearly_change / Opening_price)
    'Add total stock volumn
    Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
    'Print summary title row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Print sticker in the summary table
    ws.Range("I" & Summary_table_row).Value = Ticker_symbol
    'Print yearly change in the summary table
    ws.Range("J" & Summary_table_row).Value = Yearly_change
    ws.Range("J" & Summary_table_row).NumberFormat = "0.00"
    'Print percent change in the summary table
    ws.Range("K" & Summary_table_row).Value = Percent_change
    ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
    'Print total stok volume in the summary table
    ws.Range("L" & Summary_table_row).Value = Total_stock_volume
    
    'Add one to the summary table row
    Summary_table_row = Summary_table_row + 1
    'Reset the total stock volume
    Total_stock_volume = 0
    'Reset the opening price
    Opening_price = ws.Cells(i + 1, 3)
    
Else
    'Add to the total stock volume
    Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
    
End If

Next i
    
'Count the number of the row in summary table
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Formattingwith hightlight positive change in green and negative change in red
For j = 2 To lastrow

    If ws.Cells(j, 10).Value >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    
    Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
    
    End If
  Next j
   
 'Label the cells
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 
   For k = 1 To lastrow
  
  'Return the stock with the greatest % increase
  If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) Then
    ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
'Return the stock with the greatest % decrease
ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) Then
    ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
   ws.Cells(3, 17).NumberFormat = "0.00%"
'Return the stock with the greatest total volume
ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) Then
    ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
   
   End If
   
   Next k
   
   Next ws
   
End Sub


