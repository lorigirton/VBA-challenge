Sub Ticker_summary_table()

'Loop through each worksheet
For Each ws In Worksheets

'Set an intial variable for holding the ticker
Dim ticker As String

'Set an initial variable for holding the total stock volume
Dim total_stock_volume As Double
total_stock_volume = 0

'Set an initial variable for holding the yearly change
Dim yearly_change As Double
yearly_change = 0

'Set an initial variable for holding the percent change
Dim percent_change As Double
percent_change = 0

'Keep track of the location for each ticker in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

'Set an initial variable for holding the greatest % Increase
Dim greatest_increase As Double
greatest_increase = 0

'Format greatest %increase/decrease value as percent
ws.Range("Q2:Q3").NumberFormat = "0.00%"

'Create Headers
ws.Cells(1, "i").Value = "Ticker"
ws.Cells(1, "j").Value = "Yearly Change"
ws.Cells(1, "k").Value = "Percent Change"
ws.Cells(1, "l").Value = "Total Stock Volume"
ws.Cells(2, "o").Value = "Greatest % Increase"
ws.Cells(3, "o").Value = "Greatest % Decrease"
ws.Cells(4, "o").Value = "Greatest Total Volume"
ws.Cells(1, "p").Value = "Ticker"
ws.Cells(1, "q").Value = "Value"

'Define last row
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Declare Opening value
open_value = ws.Cells(2, "C")

'Set up For Loop to get total volume to post for each ticker
For i = 2 To last_row

    'Check we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Set the ticker name
        ticker = ws.Cells(i, 1).Value

        'Calculate yearly change
        yearly_change = ws.Cells(i, 6).Value - open_value
        If ws.Range("K" & summary_table_row).Value >= 0 Then
            ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
        End If
        

        'Calculate percent change
        percent_change = yearly_change / open_value

        If ws.Range("J" & summary_table_row).Value >= 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        End If
        
        'Add to the total stock volume
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

        'Print the ticker name in the Summary Table
        ws.Range("I" & summary_table_row).Value = ticker
        
        'Print yearly change in the Summary Table
        ws.Range("J" & summary_table_row).Value = yearly_change

        'Print percent change in the Summary Table
        ws.Range("K" & summary_table_row).Value = percent_change
        
        'Format percent change as percent
        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
        
        'Print the total stock volume to the Summary Table
        ws.Range("L" & summary_table_row).Value = total_stock_volume

        'Add one to the summary table row
        summary_table_row = summary_table_row + 1
        
        'Reset the yearly change
        yearly_change = 0
        
        'Reset percent change
        percent_change = 0
        
        'Reset the total stock volume
        total_stock_volume = 0

        'Reset open value
        open_value = ws.Cells(i + 1, "C").Value

    'If the cell immediately following a row is the same brand...
    Else

        'Add to the yearly change
        yearly_change = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
        
        'Add to the total stock volume
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

    End If

Next i

'Populate Greatest % Increase
ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & summary_table_row))
Position = WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K2:K" & summary_table_row), 0)
ws.Cells(2, 16).Value = ws.Cells(Position + 1, 9).Value

'Populate Greatest % Decrease
ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & summary_table_row))
Position = WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("K2:K" & summary_table_row), 0)
ws.Cells(3, 16).Value = ws.Cells(Position + 1, 9).Value

'Populate Greatest Total Volume
ws.Cells(4, 17).Value = WorksheetFunction.Min(ws.Range("L2:L" & summary_table_row))
Position = WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L2:L" & summary_table_row), 0)
ws.Cells(4, 16).Value = ws.Cells(Position + 1, 9).Value

Next ws

End Sub
