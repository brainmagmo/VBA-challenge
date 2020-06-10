Sub analyze_stocks()
    Dim i, j, max As Long
    max_years = ActiveWorkbook.Worksheets.count
    For i = 1 To max_years
        analyze_sheet((i))
        analyze_analysis((i))
    Next i
End Sub

''''Cells(Rows.Count, 1).End(xlUp).Row gives last row of first column

Function analyze_sheet(Sheet_Num As Long) As Long

    Dim count, iter As Long
    'number of tickers
    count = 0
    'row index
    iter = 2

    'stock stats
    Dim open_price_start, open_price_end As Double
    open_price_start = 0
    open_price_end = 0

    Dim stock_volume, stock_delta As Double
    stock_volume = 0

    'grab stock name
    cur_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1).Value

    'get starting price, day1 volume
    If cur_ticker <> "" Then
        count = 1
        open_price_start = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 3).Value
        stock_volume = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 7).Value

        'write new column titles
        ActiveWorkbook.Worksheets(Sheet_Num).Cells(1, 9) = "Ticker"
        ActiveWorkbook.Worksheets(Sheet_Num).Cells(1, 10) = "Yearly Change"
        ActiveWorkbook.Worksheets(Sheet_Num).Cells(1, 11) = "Percent Change"
        ActiveWorkbook.Worksheets(Sheet_Num).Cells(1, 12) = "Total Volume"
    End If

    ''''nex_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1).Value
    'For each row of data,
    For iter = iter To ActiveWorkbook.Worksheets(Sheet_Num).Cells(Rows.count, 1).End(xlUp).Row
        'if we have reached the end of this stock's data
        If ActiveWorkbook.Worksheets(Sheet_Num).Cells((iter + 1), 1) <> cur_ticker Then

            'grab final opening price
            open_price_end = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 3).Value

            'record the stock name, raw difference in yearly price and volume
            ActiveWorkbook.Worksheets(Sheet_Num).Cells(count + 1, 9).Value = cur_ticker
            ActiveWorkbook.Worksheets(Sheet_Num).Cells(count + 1, 10).Value = open_price_end - open_price_start
            If open_price_start <> 0 Then
                ActiveWorkbook.Worksheets(Sheet_Num).Cells(count + 1, 11).Value = (open_price_end - open_price_start) / open_price_start
            Else
                ActiveWorkbook.Worksheets(Sheet_Num).Cells(count + 1, 11).Value = ""
            End If
            ActiveWorkbook.Worksheets(Sheet_Num).Cells(count + 1, 12).Value = stock_volume

            'grab next stock's data
            open_price_start = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter + 1, 3).Value
            cur_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter + 1, 1)
            stock_volume = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter + 1, 7).Value

            count = count + 1
        Else
            'if there's more data, we scrape the volume
            stock_volume = stock_volume + ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 7).Value
        End If
    Next iter

    '2 = white
    '3 = red
    '4 = green

    For iter = 2 To count
        If ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value < 0 Then
            ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Interior.ColorIndex = 3
        ElseIf ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value > 0 Then
            ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Interior.ColorIndex = 4
        ElseIf ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value = 0 Or ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value = "" Then
            ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Interior.ColorIndex = 2
        End If
    Next iter



    'Do While cur_ticker <> ""
    '    If ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1) <> cur_ticker Then
    '        cur_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1)
    '        count = count + 1
    '    End If
    '
    '    iter = iter + 1
    'Loop

    analyze_sheet = count
End Function


Function analyze_analysis(Sheet_Num As Long) As Long
    Dim max_pos, max_neg, max_vol, iter As Long
    max_pos = 0#
    max_neg = 0#
    max_vol = 0#

    Dim pos_tkr, neg_tkr, vol_tkr As String

    For iter = 2 To ActiveWorkbook.Worksheets(Sheet_Num).Cells(Rows.count, 11).End(xlUp).Row
        ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).NumberFormat = "0.00%"
        If ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value > max_pos Then
            max_pos = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value
            pos_tkr = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 9).Value
        ElseIf ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value < max_neg Then
            max_neg = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 11).Value
            neg_tkr = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 9).Value
        End If

        If ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 12) > max_vol Then
            max_vol = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 12).Value
            vol_tkr = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 9).Value
        End If

    Next iter

    ActiveWorkbook.Worksheets(Sheet_Num).Cells(2, 14) = "Greatest % Increase"
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(3, 14) = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(4, 14) = "Greatest Volume"

    ActiveWorkbook.Worksheets(Sheet_Num).Cells(2, 15) = pos_tkr
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(3, 15) = neg_tkr
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(4, 15) = vol_tkr

    ActiveWorkbook.Worksheets(Sheet_Num).Cells(2, 16) = max_pos
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(2, 16).NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(3, 16) = max_neg
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(3, 16).NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(Sheet_Num).Cells(4, 16) = max_vol


    analyze_analysis = 1
End Function

' VBA-challenge
'VBA 'Bootcamp 'Stocks

'Todo: make VBA files
'VBA1 should:
' loop through all of the stocks of a single year

'fill a page with the following for each stock:
'  ticker_symbol
' delta_price = Delta( opening price )/d1year
'delta_percent = delprice / openingprice_0
' total stock volume over the year

'adorn the raw changes:
' highlight positive change in green, negative change in red.


'VBA2 shoudl:
' return the stock with the
'  "Greatest % increase",
' "Greatest % decrease"
'"Greatest total volume".

' loop over each sheet


'Images should:
' A screen shot for each year of your results on the Multi Year Stock Data

'Todo: clone repository, add files locally, push respository with files


