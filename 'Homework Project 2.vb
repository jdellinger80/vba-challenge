'Homework Project 2
' Create a script that loops through all the stocks for one year and outputs the following information:
'  The ticker symbol
'  Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'  The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'  The total Volume of the stock
' Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
'Bonus
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

Sub tickloop()
    

        'Setting a variable for holding the ticker name
        Dim tickname As String
    
        'Setting up a varable for holding a total count on the total volume of trade
        Dim tickvolume As Double
        tickvolume = 0

        'Keeping track of the location for each ticker name in the summary
        Dim sum_tick_row As Integer
        sum_tick_row = 2
        
        'Yearly Change, Close Price at the end of a trading year minus Open Price at the start of the trading year
        'Percent change  ((Close Price - Open Price)/Open Price)*100
        Dim open_price As Double
        'Setup of initial market open price. Remaining opening prices will be determined from a conditional loop.
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim year_change As Double
        Dim percnt_change As Double

        'Summary Table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Counting up the number of rows in our first column
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Now Loop through the rows by the stock tickers name
        'Make sure that the stock tickers are sorted and are string variables.
        
        For i = 2 To lastrow

            'Searching for where the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Establish ticker name
              tickname = Cells(i, 1).Value

              'Add the volume of trade
              tickvolume = tickvolume + Cells(i, 7).Value

              'Print the stock ticker name in the summary
              Range("I" & sum_tick_row).Value = tickname

              'Print the trading volume for each ticker in the summary
              Range("L" & sum_tick_row).Value = tickvolume

              'Find closing price
              close_price = Cells(i, 6).Value

              'Tabulate yearly change
               year_change = (close_price - open_price)
              
              'Print the yearly change for each stock ticker represented
              Range("J" & sum_tick_row).Value = year_change

              'Find any potential non-divisibile Zeros, when calculating the data sets percent change
                If open_price = 0 Then
                    percnt_change = 0
                Else
                    percnt_change = year_change / open_price
                End If

              'Print out the yearly change for each ticker
              Range("K" & sum_tick_row).Value = percnt_change
              Range("K" & sum_tick_row).NumberFormat = "0.00%"
   
              'Reset the counter. Add one to the sum_tick_row
              sum_tick_row = sum_tick_row + 1

              'Reset trading volumes back to zero
              tickvolume = 0

              'Reset the opening price
              open_price = Cells(i + 1, 3)
            
            Else
              
               'Add trading volumes
              tickvolume = tickvolume + Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting highlights positive change in green and negative change in red
    'Find the last row of the summary table

    lastrow_summry_tab = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summry_tab
            If Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
            Else
                Cells(i, 10).Interior.ColorIndex = 10
            End If
        Next i

    'Highlight the stock price changes
 

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    'Determine max and min values in column "Percent Change" and max in column "Total Stock Volume"
    'Collect stock ticker and the values for it's percent change and total volume of trade for stock tickers
    
        For i = 2 To lastrow_summry_tab
            'Find the max percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summry_tab)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Find the min percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summry_tab)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the max volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summry_tab)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub '
