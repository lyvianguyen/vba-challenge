Sub stockdata()
    
    'define everything
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume As Double
    Dim summarytable As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim price As Double
    Dim annual_change As Double
    Dim percent_change As Double
    Dim lastRow As Long
    For Each ws In Worksheets
    
    'define the variables
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As Double
    
    'set values for variables
    volume = 0
    summarytable = 2
    price = 2
    percent_change = 0
    greatestincrease = 0
    greatestdecrease = 0
    greatestvolume = 0
    
    'set headers for spreadsheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest Percent Increase"
    ws.Cells(3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'define the last row - referencing syntax from day 3 learning vba activity credit card
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'loop through all stocks and pull unique values
    For i = 2 To lastRow
    
        'add these values towards total volume for specific ticker
        volume = volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
        
        'pull ticker and volume info into summary table
        ws.Range("I" & summarytable).Value = ticker
        
        ws.Range("L" & summarytable).Value = volume
        
        'reset stock volume to 0 at the end of each ticker
        
        volume = 0
        
        'Set values of open close and annual change
        open_price = ws.Range("C" & price)
        close_price = ws.Range("F" & i)
        
        annual_change = close_price - open_price
        
        ws.Range("J" & summarytable).Value = annual_change
        
            'Formula for percent change calculation
            If open_price <> 0 Then
            open_price = ws.Range("C" & price)
            percent_change = (close_price - open_price) / open_price
            
            End If
        
            'Show percent change in summary data and format as percentage - https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
        
            ws.Range("K" & summarytable).Value = percent_change
            ws.Range("K" & summarytable).NumberFormat = "0.00%"
        
            'reset values
        
            percent_change = 0
            summarytable = summarytable + 1
            price = i + 1
        
            End If
        
    Next i
    
    'define the last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set cell values and pull data - assistance from vol data director
    
    For i = 2 To lastRow
    
    If ws.Range("K" & i).Value > ws.Cells(2, 17).Value Then
    ws.Range("Q2").Value = ws.Range("K" & i).Value
    ws.Range("P2").Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("K" & i).Value < ws.Cells(3, 17).Value Then
    ws.Range("Q3").Value = ws.Range("K" & i).Value
    ws.Range("P3").Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("L" & i).Value > ws.Cells(4, 17).Value Then
    ws.Range("Q4").Value = ws.Range("L" & i).Value
    ws.Range("P4").Value = ws.Range("I" & i).Value
    
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    End If
    
    Next i

 'Set cell color - red for negative green for positive - class activity
        
    If ws.Range("J" & summarytable).Value > 0 Then
        ws.Range("J" & summarytable).Interior.ColorIndex = 4
    Else
        ws.Range("J" & summarytable).Interior.ColorIndex = 3
    End If
            
    Next ws
    
End Sub