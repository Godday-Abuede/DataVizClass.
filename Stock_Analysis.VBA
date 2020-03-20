Sub Stock_Analysis()
    
    'Create a variable to hold the worksheet and also the output message
    Dim ws As Worksheet

    'Loop through each worksheet in the Worksheets collection of the active workbook
    For Each ws In Worksheets

        'column labels for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'variable to hold the ticker symbol
        Dim ticker_symbol As String

        'variable to hold total volume of stock traded
        Dim total_vol As Double
        total_vol = 0

        'tracker of location for each ticker symbol in the summary table
        Dim rowcount As Long
        rowcount = 2

        'variable to hold year open price
        Dim year_open As Double
        year_open = 0

        'variable to hold year close price
        Dim year_close As Double
        year_close = 0
        
        'variable to hold the change in price for the year
        Dim year_change As Double
        year_change = 0

        'variable to hold the percent change in price for the year
        Dim percent_change As Double
        percent_change = 0

        'variable for total rows to loop through
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'loop through ticker symbols
        For i = 2 To lastrow
            
            'condition to get year open price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If

            'sum the volume for each row to determine total stock volume for the year
            total_vol = total_vol + ws.Cells(i, 7)

            'condition to determine changes in the ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'ticker symbol to summary table
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                'total stock volume to the summary table
                ws.Cells(rowcount, 12).Value = total_vol

                'get year end price
                year_close = ws.Cells(i, 6).Value

                'compute price change for the year and move to the summary table.
                year_change = year_close - year_open
                ws.Cells(rowcount, 10).Value = year_change

                'highlight positive or negative change.
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'compute percent change for the year and move to summary table as a percentage.
                If year_open = 0 And year_close = 0 Then
                    'starts at zero and ends at zero,i.e zero increase.  due to zero divisor, no formular.
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf year_open = 0 Then
                    'evaluate actual price increase and then, have new stock as percent change.
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'addition of 1 to rowcount in order to move it to the next empty row in the summary table
                rowcount = rowcount + 1

                'Reset total stock volume, year open and clos prices, year change, and year percent change
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        Next i

        ' best and worst performance table
        'Titles
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'the lastrow to count the number of rows in the summary table
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Set variables to hold best and worst performers, and stock with most volume
        Dim best_stock As String
        Dim best_value As Double

        'make best performer equals first stock
        best_value = ws.Cells(2, 11).Value

        Dim worst_stock As String
        Dim worst_value As Double

        'make worst performer equals first stock
        worst_value = ws.Cells(2, 11).Value

        Dim most_vol_stock As String
        Dim most_vol_value As Double

        'make most volume equals first stock
        most_vol_value = ws.Cells(2, 12).Value

        'loop through summary table
        For j = 2 To lastrow

            'determine best performer
            If ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_stock = ws.Cells(j, 9).Value
            End If

            'determine worst performer
            If ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_stock = ws.Cells(j, 9).Value
            End If

            'determine stock with greatest volume traded
            If ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_stock = ws.Cells(j, 9).Value
            End If

        Next j

        'move the best and worst performers, and stock with most volume items to performance table
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value

        'Autofit table columns
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws

End Sub
