Attribute VB_Name = "Module1"
Sub StockMarketSummary():

For Each ws In Worksheets
    
    'Labeling the Additional Columns
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Establishing Variables
    
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVolume As Double
    Dim tableRow As Integer
    Dim row As Double
    Dim lastRow As Double
    
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    tableRow = 2
    
    'Looping through each row until the last row
    
    For row = 2 To lastRow
    
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        
            'Printing the ticker code
            
            ticker = ws.Cells(row, 1).Value
            ws.Cells(tableRow, 9).Value = ticker
      
            'Getting the yearly change and printing
            
            yearlyChange = closingPrice - openingPrice
            ws.Cells(tableRow, 10).Value = yearlyChange
            
            'Highlight Postive Changes as Green and Negative Changes as Red
                If yearlyChange >= 0 Then
                
                    ws.Cells(tableRow, 10).Interior.ColorIndex = 4
                    
                Else
                
                    ws.Cells(tableRow, 10).Interior.ColorIndex = 3
            
                End If
                
               'Getting the percent change and printing
               
                If openingPrice <> 0 Then
                
                    percentChange = (yearlyChange / openingPrice) * 100
                    ws.Cells(tableRow, 11).Style = "Percent"
                    ws.Cells(tableRow, 11).Value = percentChange
                    
                Else
                
                    percentChange = 100
                    ws.Cells(tableRow, 11).Style = "Percent"
                    ws.Cells(tableRow, 11).Value = percentChange
                    
                End If
                
            'Printing total stock volume
            
            ws.Cells(tableRow, 12).Value = totalStockVolume
            
            'Zero-ing out totals
            
            openingPrice = 0
            closingPrice = 0
            totalStockVolume = 0
            
            'Moving down to the next row
            
            tableRow = tableRow + 1
        
        Else
        
            'Adding to the Opening and Closing Price
            
            openingPrice = openingPrice + ws.Cells(row, 3).Value
            closingPrice = closingPrice + ws.Cells(row, 6).Value
            
            'Adding to the Total Stock Volume
            
            totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value
        
        End If
    
    Next row
    
    'BONUS
    
    'Labeling the Additional Columns
    
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increased"
    ws.Cells(3, 14).Value = "Greatest % Decreased"
    ws.Cells(4, 14).Value = "Greatest Total Value"
    
    'Establishing Variables
    
    Dim tableLastRow As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    tableLastRow = Cells(Rows.Count, 9).End(xlUp).row
    
    'Starting variable amounts to 0
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    'Looping through the table to get the greatest values
    
    For row = 2 To tableLastRow
        
        'Getting the Greatest % Increase
        
        If ws.Cells(row, 11) > greatestIncrease Then
        
            greatestIncrease = ws.Cells(row, 11)
            tickerIncrease = ws.Cells(row, 9)
        
        End If
        
        'Getting the Greatest % Decrease
        
        If ws.Cells(row, 11) < greatestDecrease Then
        
            greatestDecrease = ws.Cells(row, 11)
            tickerDecrease = ws.Cells(row, 9)
            
        End If
    
        'Getting the Greatest Total Volume
        
        If ws.Cells(row, 12) > greatestVolume Then
        
            greatestVolume = ws.Cells(row, 12)
            tickerVolume = ws.Cells(row, 9)
            
        End If
    
       Next row

    'Printing the greatest values
    
         ws.Cells(2, 15).Value = tickerIncrease
         ws.Cells(2, 16).Value = greatestIncrease
         ws.Cells(2, 16).Style = "Percent"
        
        ws.Cells(3, 15).Value = tickerDecrease
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(3, 16).Style = "Percent"
        
        ws.Cells(4, 15).Value = tickerVolume
        ws.Cells(4, 16).Value = greatestVolume
        
    ws.Cells.EntireColumn.AutoFit

Next ws

End Sub
