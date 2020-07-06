Sub stock()

'Set Dimensions for variables
Dim i, summary_table_row As Integer
Dim stock_volume, year_close, year_open, yearly_change, percent_change, greatest_increase, greatest_decrease, greatest_volume As Double
Dim ticker_name, ticker1, ticker2, ticker3 As String
Dim ws As Worksheet

'Apply code to each worksheet/a loop

    For Each ws In Worksheets

'Set counters. One for rows in summary table and the other total stock volume
'Set year open variable to capture first open price before loop

        summary_table_row = 2
        stock_volume = 0
        year_open = ws.Cells(2, 3).Value

'Set titles to summary table

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"


'Auto calculate last rows for loops

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Start loop. Use conditional to ID different ticker
'ID variables needed for readability
'Move table row after input and reset stock volume.
'first open is previously defined, close will always be the defined variable, open will be redefined at the end
'If condition is not met, the ticker is the same; add to stock volume count
'First if statement is to account for 0 errors

            For i = 2 To lastrow
    
                If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) And year_open = 0 Then
                
                    ticker_name = ws.Cells(i, 1).Value
                    year_close = 0
                    yearly_change = 0
                    percent_change = 0
                    ws.Cells(summary_table_row, 9).Value = ticker_name
                    ws.Cells(summary_table_row, 10).Value = yearly_change
                    ws.Cells(summary_table_row, 11).Value = percent_change
                    ws.Cells(summary_table_row, 12).Value = stock_volume + ws.Cells(i, 7).Value
                    summary_table_row = summary_table_row + 1
                    stock_volume = 0
                    year_open = ws.Cells(i + 1, 3).Value
    
                ElseIf (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            
                    ticker_name = ws.Cells(i, 1).Value
                    year_close = ws.Cells(i, 6).Value
                    yearly_change = year_close - year_open
                    percent_change = (yearly_change / year_open)
                    ws.Cells(summary_table_row, 9).Value = ticker_name
                    ws.Cells(summary_table_row, 10).Value = yearly_change
                    ws.Cells(summary_table_row, 11).Value = percent_change
                    ws.Cells(summary_table_row, 12).Value = stock_volume + ws.Cells(i, 7).Value
                    summary_table_row = summary_table_row + 1
                    stock_volume = 0
                    year_open = ws.Cells(i + 1, 3).Value
            
                ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
        
                    stock_volume = stock_volume + ws.Cells(i, 7).Value
            
                End If
        
            Next i

'we need to set a last row count for unique IDS
'This loop has to do with formatting after the primary loop is ran
'if the diffence of yearly open to close is more than 0 green, if less than 0 red


        lastrowv2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
            For i = 2 To lastrowv2
        
                If (ws.Cells(i, 10).Value) > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                
                ElseIf (ws.Cells(i, 10).Value) < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                
                ElseIf (ws.Cells(i, 10).Value) = 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 0
                
                End If
            
            Next i
            
'Format % change column

       ws.Range("K2:K" & lastrowv2).NumberFormat = "0.00%"
        
'title,ID, and populate new table (titles in tables reveal purpose)
        
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

        greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & lastrowv2))
        greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & lastrowv2))
        greatest_volume = WorksheetFunction.Max(ws.Range("L2:L" & lastrowv2))

        ws.Range("P2").Value = greatest_increase
        ws.Range("P3").Value = greatest_decrease
        ws.Range("P4").Value = greatest_volume
        
        ws.Range("P2:P3").NumberFormat = "0.00%"

'Match the greatest numbers to its ticker. Plus 1 is needed to adjust for starting at K2 (has to do with lastrowv2 formula)

        ticker1 = ws.Cells(WorksheetFunction.Match(greatest_increase, ws.Range("K2:K" & lastrowv2), 0) + 1, 9)
        ticker2 = ws.Cells(WorksheetFunction.Match(greatest_decrease, ws.Range("K2:K" & lastrowv2), 0) + 1, 9)
        ticker3 = ws.Cells(WorksheetFunction.Match(greatest_volume, ws.Range("L2:L" & lastrowv2), 0) + 1, 9)
        ws.Range("O2").Value = ticker1
        ws.Range("O3").Value = ticker2
        ws.Range("O4").Value = ticker3

'Autofit Column width

        ws.Columns("J:L").AutoFit
        ws.Columns("N:P").AutoFit

    Next ws
    
End Sub
