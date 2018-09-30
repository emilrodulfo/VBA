Sub VBAofWallStreet()
    
    '[MODERATE]
    
    ' Set initial variable to hold Volume of Stock
    Dim StockVolume As Double
    StockVolume = 0
    
    ' Keeps track of what row to insert new line
    Dim NewRow As Long
    NewRow = 2
    
    Dim OpeningPrice As Double

    ' Set variable for Percent Change
    Dim PercentChange As Double

    ' Find last row
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Insert Ticker column title
    Cells(1, 9).Value = "Ticker"

    ' Insert Ticker column title
    Cells(1, 10).Value = "Yearly"

    ' Insert Ticker column title
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 11).ColumnWidth = 14
    
    ' Insert Ticker column title
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 12).ColumnWidth = 16
    
    ' Loop through all tickers
    For i = 2 To LastRow
    
        ' Check that the next ticker is still the same
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Set ticker name
        TickerName = Cells(i, 1).Value
        
        ' Find opening price
        OpeningPrice = Cells(i - OpeningPriceRow, 3).Value
        
        ' Calculate Yearly Change
        YearlyChange = Cells(i, 6).Value - Cells(i - OpeningPriceRow, 3).Value
        
        ' Add to the Total Stock Volume
        StockVolume = StockVolume + Cells(i, 7).Value
        
            ' Calculate Percent Change
            If YearlyChange = 0 Or (YearlyChange - Cells(i, 6).Value) = 0 Then
            PercentChange = 0
            
            Else
            PercentChange = -YearlyChange / (YearlyChange - Cells(i, 6).Value)
            
            End If
        
        ' Print Ticker
        Range("I" & NewRow).Value = TickerName
        
        ' Print opening and closing difference
        Range("J" & NewRow).Value = YearlyChange
        
            'Add conditional formatting for interior color on Yearly Change
            If YearlyChange > 0 Then
            Range("J" & NewRow).Interior.ColorIndex = 4
            
            ElseIf YearlyChange < 0 Then
            Range("J" & NewRow).Interior.ColorIndex = 3
            
            Else
            Range("J" & NewRow).Interior.ColorIndex = 2
            
            End If
                
        ' Print opening and closing difference
        Range("K" & NewRow).Value = PercentChange
        
        ' Format the Percent Change to be a percentage with 2 decimal places
        Range("K" & NewRow).NumberFormat = "0.00%"
        
        'Print Total Stock Volume
        Range("L" & NewRow).Value = StockVolume
        
        'Reset StockVolume
        StockVolume = 0
        
        'Reset OpeningPriceRow
        OpeningPriceRow = 0
        
        'Add one to NewRow
        NewRow = NewRow + 1
        
        Else
        
        ' Add to Stock Volume
        StockVolume = StockVolume + Cells(i, 7).Value
        
        'Add to Opening Price Row
        OpeningPriceRow = OpeningPriceRow + 1
        
        End If
        
    Next i
    
    '[HARD]
    
    Dim HardLastRow As Long
    HardLastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Insert Ticker column title
    Cells(1, 16).Value = "Ticker"

    ' Insert Value column title
    Cells(1, 17).Value = "Value"

    ' Insert Greatest % Inc column title
    Cells(2, 15).Value = "Greatest % Increase"
    
    ' Insert Greatest % Dec column title
    Cells(3, 15).Value = "Greatest % Decrease"
    
    ' Insert Greatest Total Vol title
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 15).ColumnWidth = 19
    
    ' Loop through all tickers to find greatest %'s
    For i = 2 To HardLastRow
        
        If Cells(i + 1, 11).Value > GreatestInc Then
        
        GreatestInc = Cells(i + 1, 11).Value
        
        IncTicker = Cells(i + 1, 9).Value
        
        ElseIf Cells(i + 1, 11).Value < GreatestDec Then
        
        GreatestDec = Cells(i + 1, 11).Value
        
        DecTicker = Cells(i + 1, 9).Value
        
        End If
    
    Next i
    
    ' Print Greatest % Increase
    Cells(2, 16).Value = IncTicker
    Cells(2, 17).Value = GreatestInc
    Cells(2, 17).NumberFormat = "0.00%"
    
    ' Print Greatest % Decrease
    Cells(3, 16).Value = DecTicker
    Cells(3, 17).Value = GreatestDec
    Cells(3, 17).NumberFormat = "0.00%"
    
    ' Loop through all tickers to find greatest vol
    For i = 2 To HardLastRow
        
        If Cells(i + 1, 12).Value > GreatestVol Then
        
        GreatestVol = Cells(i + 1, 12).Value
        
        VolTicker = Cells(i + 1, 9).Value
        
        End If
    
    Next i
    
    ' Print Greatest % Increase
    Cells(4, 16).Value = VolTicker
    Cells(4, 17).Value = GreatestVol
    Cells(4, 17).ColumnWidth = 14
    
End Sub

