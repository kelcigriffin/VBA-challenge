'Instructions
'Create a script that loops through all the stocks for one year and outputs
'the following information:
    '1. The ticker symbol
    '2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    '3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    '4. The total stock volume of the stock.
    '5. Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    '6. Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
    '7. Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

Sub multiple_year_stock_data():
    
' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
' --------------------------------------------
    For Each ws In Worksheets
' --------------------------------------------

' --------------------------------------------
    ' Determine the last row
' --------------------------------------------
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
' --------------------------------------------
       
' Set initial variables for holding the series of ticker names, provided in
'Column "A". Define variables for first open price, last close
'price, and determine that the ticker summary starts on row 2
    Dim ticker_name As String
    Dim first_open As Double
    Dim last_close As Double
    Dim ticker_summary As Integer
  'Define variables for yearly change, percent change, and total volume
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
  
  'specify location of the first opening price and the row ticker summaries
  'begin populating
  
    first_open = ws.Cells(2, 3).Value
    ticker_summary = 2
    total_volume = 0
    
' --------------------------------------------
        'SET ALL COLUMN HEADERS
' --------------------------------------------
  'Set column headers for data summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
     
    'set column headers for final observations
        ws.Cells(1, 15).Value = " "
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

' --------------------------------------------
        'BEGIN LOOP
' --------------------------------------------
        
    ' Loop through all open and close prices
    For i = 2 To last_row
  
        'define formula for total volume before the If statement
        total_volume = total_volume + ws.Cells(i, 7).Value
    
        ' Check if we are still within the same ticker, if we are not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'give the name of the ticker being summarized
            ticker_name = ws.Cells(i, 1).Value
            ws.Cells(ticker_summary, 9).Value = ticker_name
        
            last_close = ws.Cells(i, 6).Value
           
            'calculate yearly change and add value to cells in Yearly Change column
            yearly_change = (last_close - first_open)
            ws.Cells(ticker_summary, 10).Value = yearly_change
        
            'calculate the percent change and add value to Percent Change column
            percent_change = ((last_close - first_open) / first_open)
            ws.Cells(ticker_summary, 11).Value = percent_change
            ws.Cells(ticker_summary, 11).NumberFormat = "0.00%"
        
             'calculate total volume (Column G) and add value to Total Stock Volume column
            ws.Cells(ticker_summary, 12).Value = total_volume
        
        
            'command loop to move to the next range of ticker names (Col A),
            'reset yearly change to zero, and move down a row to gather the
            'opening price (Col C) for the new ticker in the series
            'reset total volume
            ticker_summary = ticker_summary + 1

            total_volume = 0
        
            yearly_change = 0
        
            first_open = ws.Cells(i + 1, 3).Value
        
       End If
    
    Next i
    
    ' --------------------------------------------
            'CONDITIONAL STATEMENTS
    ' --------------------------------------------
    
    'setting the stage for the conditional statement, where the formula observes all rows in the Yearly Change column to determine if they're positive or negative
    last_row_yearly_change = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To last_row_yearly_change

            'change cells in Yearly Change column to red or green, based on positive or negative growth
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
                
            ElseIf ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
                               
            End If
    Next i
  
 ' --------------------------------------------
            'MIN/MAX
 ' --------------------------------------------
            
  'use built in VBA functions to find the minimum and maximum percent change from Column K, and the max
  'total stock volume from Column L. Use that data to fill in final observations, with ticker name and values listed
  'for each parameter
  For i = 2 To last_row_yearly_change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row_yearly_change)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            
            
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min((ws.Range("K2:K" & last_row_yearly_change))) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row_yearly_change)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
        End If
           
        
    Next i

    ws.Columns("A:Q").AutoFit

  Next ws

End Sub

