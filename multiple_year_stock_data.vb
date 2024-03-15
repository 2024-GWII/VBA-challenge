Sub annual_stock_stats()


Dim ticker As String                                'ticker
Dim year_open As Double                       'first opening price
Dim year_close As Double                       'last closing price
Dim year_change As Double                    'year_close - year_open
Dim percent_change As Double               'year_change / year_open
Dim year_volume As Double                   'sum of vol over ticker
Dim summary_table_row As Integer        'index of current row in the summary table
Dim ws As Worksheet                             'curent worksheet
Dim first_data_table_row As Integer       'all data talbes have first_data_table_row = 2; avoid magic number
Dim max_change As Double
Dim min_change As Double
Dim max_vol As Double

first_data_table_row = 2                        'data in data table beings in row 2; row 1 is the header row


'Loop thru all of the worksheets
For Each ws In Worksheets

    '-------------Build the first summary table-------------

    'Clear Formats
    ws.Range("I1:R1000000").ClearFormats

    'Title the columns of the first summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Start loops at 2nd row, right after the header row
    summary_table_row = 2
    ticker_first_row = 2
    year_volume = 0

    last_data_table_row = ws.Cells(Rows.Count, "A").End(xlUp).Row     'index of the last row in the data table

        'Loop thru data table, calculate stats for each ticker
        For current_data_table_row = first_data_table_row To last_data_table_row

            'Scan thru the rows of the table until current_data_table_row = the last row for a given ticker,
            'ie, the last row before the ticker changes.
            If ws.Cells(current_data_table_row + 1, 1).Value <> ws.Cells(current_data_table_row, 1).Value Then

                    'Get the ticker
                    ticker = ws.Cells(current_data_table_row, 1).Value
        
                    ' Find opening value from first row where the ticker is listed.  (Opening value will be in column #3.)
                    year_open = ws.Cells(ticker_first_row, 3).Value
                    
                    'Find closing value from the last row where the ticker is listed.  (Closing value will be in column #6.)
                    year_close = ws.Cells(current_data_table_row, 6).Value
        
                    ' Loop over all of the rows for a given ticker and sum the volumes, listed in column #7.
                    For ticker_row = ticker_first_row To current_data_table_row
                        year_volume = year_volume + ws.Cells(ticker_row, 7).Value
                    Next ticker_row
        
                    'Calculate changes
                    year_change = year_close - year_open
                    percent_change = year_change / year_open
        
                    'Put the stats in the summary table...
                    ws.Cells(summary_table_row, 9).Value = ticker
                    ws.Cells(summary_table_row, 10).Value = year_change
                    ws.Cells(summary_table_row, 11).Value = percent_change
                    ws.Cells(summary_table_row, 12).Value = year_volume
        
                    'Initialize for next ticker...
                    summary_table_row = summary_table_row + 1
                    year_volume = 0
                    year_change = 0
                    percent_change = 0
        
                    '...and move to the next ticker
                    ticker_first_row = current_data_table_row + 1

            End If 'End the conditional looking for the row where ticker changes

    Next current_data_table_row 'Continue advancing through the data table to find the next ticker change


    '-------------Build the second summary table------------

    'Labels in the second summary table
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    'Use the Excel Min and Max Functions to pull out the data, rather than looping.
    Set percent_range = ws.Range("K:K") 'Column with the percents
    Set vol_range = ws.Range("L:L") 'Column with the volumes

    'Find min and max of percents and volumes
    max_percent = Application.WorksheetFunction.Max(percent_range)
    min_percent = Application.WorksheetFunction.Min(percent_range)
    max_vol = Application.WorksheetFunction.Max(vol_range)
    
    'Pull the tickers for min_percent, max_percent, and max_vol and put them in the table. (Tickers are in Column #9)
    ws.Range("O2").Value = ws.Cells(percent_range.Find(What:=max_percent).Row, 9).Value
    ws.Range("O3").Value = ws.Cells(percent_range.Find(What:=min_percent).Row, 9).Value
    ws.Range("O4").Value = ws.Cells(vol_range.Find(What:=max_vol).Row, 9).Value

    'And now put the min's and max's in the table
    ws.Range("P2").Value = max_percent
    ws.Range("P3").Value = min_percent
    ws.Range("P4").Value = max_vol

    'Format the %'s accordingly
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("K:K").NumberFormat = "0.00%"

    'And loop thru the rows in Column "J" / Column #10 and color them according to integer sign
    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    For j = 2 To jEndRow
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4 'Color Green if >0
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3 'Color Red if <=0
            End If
    Next j


Next ws         'Loop to next worksheet


End Sub
