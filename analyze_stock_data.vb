Sub manage_sheets()
    Dim i, j, max As Long
    max_years = ActiveWorkbook.Worksheets.count
    For i = 1 To max_years
        'Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count)).Name = ActiveWorkbook.Worksheets(i).Name + "_Analysis"
        For j = 2 To (get_ticker_count((i)) + 1)
            'Cells(j, 2) = 0
        
        Next j
    Next i
End Sub

'Cells(Rows.Count, 1).End(xlUp).Row gives last row of first column

Function get_ticker_count(Sheet_Num As Long) As Long

    Dim count, iter As Long
    count = 0
    iter = 2
    
    Dim open_price_start, open_price_end As Double
    open_price_start = 0
    open_price_end = 0
    
    Dim stock_volume As Long
    stock_volume = 0
    
    cur_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1).Value
    
    If cur_ticker <> "" Then
        count = 1
        iter = iter + 1
        open_price_start = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 7).Value
        stock_volume = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 7).Value
        
    End If
    
    'nex_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1).Value
    
    For iter = iter To Cells(Rows.count, 1).End(xlUp).Row
        If ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1) <> cur_ticker Then
            cur_ticker = ActiveWorkbook.Worksheets(Sheet_Num).Cells(iter, 1)
            count = count + 1
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
    
    get_ticker_count = count
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


