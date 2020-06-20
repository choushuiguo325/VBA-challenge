sub tracker():
dim ticker as string
dim opening, closing, yearly_change,percent_change,max_increase,min_increase as double
dim total,ticker_day,ticker_count,incre_position,decre_position,volume_position as long


' create a loop through each worksheet
for each ws in worksheets
    
    ticker_day = 0
    ticker_count = 0
    total = 0
    max_increase = 0
    min_increase = 0
    max_volume = 0

    ws.range("I1").value = "Ticker"
    ws.range("J1").value = "Yearly Change"
    ws.range("K1").value = "Percent Change"
    ws.range("L1").value = "Total Stock Volume"

    ws.range("O2").value = "Greatest % Increase"
    ws.range("O3").value = "Greatest % Decrease"
    ws.range("O4").value = "Greatest Total Volume"
    ws.range("P1").value = "Ticker"
    ws.range("Q1").value = "Value"
    last_row = ws.cells(rows.count,1).end(xlup).row

    ' create a loop through each row in each worksheet
    for i = 2 to last_row
        'create loops to summarize the tickers
        'if the ticker is equal to the next ticker
        if ws.range("A"&i).value = ws.range("A"&(i+1)).value then
            'sum the total stock volume
            total = total + ws.range("G"&i).value
            ticker_day = ticker_day + 1
        'if not
        elseif ws.range("A"&i).value <> ws.range("A"&(i+1)).value then
        'store the brand, opening price and closing price of the last ticker
            ticker_count = ticker_count + 1
            ticker = ws.range("A"&i).value
            opening = ws.range("C"&(i-ticker_day)).value
            closing = ws.range("F"&i).value

            'calculate the yearly change and percent change, using "if" to fill the interior color
            yearly_change = closing - opening
            ws.range("J"&(ticker_count+1)).value = yearly_change
            
            if yearly_change < 0 then  
                ws.range("J"&(ticker_count+1)).interior.colorindex = 3
            else
                ws.range("J"&(ticker_count+1)).interior.colorindex = 4
            end if 

            'judge the value of denominator
            if opening = 0 then
                ws.range("K"&(ticker_count+1)).value = "None"
            else
                percent_change = formatpercent(yearly_change/opening)
                ws.range("K"&(ticker_count+1)).value = percent_change
            end if 

            'calculate the final total stock volume
            total =  total + ws.range("G"&i).value
            ws.range("L"&(ticker_count+1)).value = total
            ws.range("I"&(ticker_count+1)).value = ticker
            total = 0
            ticker_day = 0

        end if 

    next i 

    'create loops to retrive the ticker with our interests.
    for j = 2 to (ticker_count+1)
        ' position of maximum and minimum percent change
        if ws.range("K"&j).value <> "None" then
            if ws.range("K"&j).value > max_increase then
                max_increase = ws.range("K"&j).value
                incre_position = j
            end if 
            if ws.range("K"&j).value < min_increase then
                min_increase = ws.range("K"&j).value
                decre_position = j
            end if
        end if 
        if ws.range("L"&j).value > max_volume then
            max_volume = ws.range("L"&j).value
            volume_position = j
        end if 
    next j

    'greatest increase
    ws.range("P2").value = ws.range("I"&incre_position).value
    ws.range("Q2").value = formatpercent(max_increase)
    
    'greatest decrease
    ws.range("P3").value = ws.range("I"&decre_position).value
    ws.range("Q3").value = formatpercent(min_increase)
    
    'greatest total stock volume
    ws.range("P4").value = ws.range("I"&volume_position).value
    ws.range("Q4").value = max_volume
    ws.Columns("A:Q").AutoFit
next ws

end sub