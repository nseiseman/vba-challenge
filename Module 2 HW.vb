Sub stock_analysis()

    'Put ws. in front of Range and Cells to run loop for all sheets
    Dim ws As Worksheet

    'Loop that runs through all the sheets creating the main table and summary table
    For each ws in Worksheets
        'Create main table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Define main table calculation variables
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double

        'Use main_table_row for creating the inputs to the main table
        'Initialize - main table row counter
        Dim main_table_row As Long
        'Set initial row
        main_table_row = 2

        'Used to create the list of unique stocks in column I
        'Initialize - tracking the current stock ticker
        Dim current_ticker As String
        'Set first current ticker
        current_ticker = ws.Cells(2, 1).Value

        'Used to find the open_price for a given stock
        'Initialize - tracking the first row of a ticker's data
        Dim ticker_data_first_row As Long
        'Set initial row
        ticker_data_first_row = 2

        'Create summary table headers
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

        'Define summary table variables
        Dim greatest_increase As Double
        Dim greatest_decrease As Double 
        Dim greatest_volume As Double
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_volume_ticker As String

        'Set initial values for summary table variables
        greatest_increase = 0
        greatest_decrease = 0

        'Loop for main and summary table
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            'Create unique ticker list in column I
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Find the ticker
                current_ticker = ws.Cells(i, 1).Value
            
                'Print the current ticker in the main table
                ws.Range("I" & main_table_row).Value = current_ticker
            
                'Calculate - yearly_change and percent_change
                'finds open price of first day
                'Uses ticker_data_first_row because the open price is in the first row for a given ticker
                open_price = ws.Cells(ticker_data_first_row, 3).Value
                'finds close price of last day
                'Uses i because the close_price is in the last row before the new ticker
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                percent_change = yearly_change / open_price
            
                'Calculate - total_stock_volume
                total_stock_volume = 0
                For j = ticker_data_first_row To i
                    total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
                Next j
            
                'Print the calculated values in the table
                ws.Range("J" & main_table_row).Value = yearly_change
                ws.Range("K" & main_table_row).Value = percent_change * 100 & "%"
                ws.Range("L" & main_table_row).Value = total_stock_volume
            
                'Calculate and store the greatest increase and greatest decrease values
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = current_ticker
                ElseIf percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = current_ticker
                End If

                'Calculate and store the greats stock volume values
                If total_stock_volume > greatest_volume Then
                    greatest_volume = total_stock_volume
                    greatest_volume_ticker = current_ticker
                End If

                'Move to the next row in the table
                'Add 1 to main_table_row to go to next ticker
                main_table_row = main_table_row + 1
                'Add 1 to i to go to next ticker in the ticker data
                ticker_data_first_row = i + 1
            
            End If
        Next i

        'Create the summary table
        ws.Range("O2").Value = greatest_increase_ticker
        ws.Range("O3").Value = greatest_decrease_ticker
        ws.Range("O4").Value = greatest_volume_ticker
        ws.Range("P2").Value = greatest_increase * 100 & "%"
        ws.Range("P3").Value = greatest_decrease * 100 & "%"
        ws.Range("P4").Value = greatest_volume

    Next ws
End Sub