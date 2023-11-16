Sub StockSummary()
    
    Dim tickerSymbol As String
    Dim yearlyOpeningPrice As Double
    Dim yearlyClosingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim row As Long
    
    Dim outputRow As Integer
    outputRow = 2
    Dim maxPercentageIncrease As Double
    Dim maxPercentageDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentageIncreaseTicker As String
    Dim maxPercentageDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    yearlyOpeningPrice = Cells(2, 3).Value '
    tickerSymbol = Cells(2, 1).Value
    
    
    For row = 2 To Cells(Rows.Count, 1).End(xlUp).row
        
        If Cells(row, 1).Value <> tickerSymbol Then
            
            yearlyClosingPrice = Cells(row - 1, 6).Value
            yearlyChange = yearlyClosingPrice - yearlyOpeningPrice
            percentageChange = yearlyChange / yearlyOpeningPrice
            totalVolume = Application.WorksheetFunction.Sum(Range(Cells(outputRow, 7), Cells(row - 1, 7)))
            
            Cells(outputRow, 9).Value = tickerSymbol
            Cells(outputRow, 10).Value = yearlyChange
            Cells(outputRow, 11).Value = percentageChange
            Cells(outputRow, 12).Value = totalVolume
            
            yearlyOpeningPrice = Cells(row, 3).Value
            tickerSymbol = Cells(row, 1).Value
            
            outputRow = outputRow + 1
        End If
    
        If row = Cells(Rows.Count, 1).End(xlUp).row Then
        
            yearlyClosingPrice = Cells(row, 6).Value
            yearlyChange = yearlyClosingPrice - yearlyOpeningPrice
            percentageChange = yearlyChange / yearlyOpeningPrice
            totalVolume = Cells(row, 7).Value
            
            Cells(outputRow, 9).Value = tickerSymbol
            Cells(outputRow, 10).Value = yearlyChange
            Cells(outputRow, 11).Value = percentageChange
            Cells(outputRow, 12).Value = totalVolume
        End If
        
        If Cells(outputRow, 11).Value > maxPercentageIncrease Then
            maxPercentageIncrease = Cells(outputRow, 11).Value
            maxPercentageIncreaseTicker = Cells(outputRow, 9).Value
        End If
        If Cells(outputRow, 11).Value < maxPercentageDecrease Then
            maxPercentageDecrease = Cells(outputRow, 11).Value
            maxPercentageDecreaseTicker = Cells(outputRow, 9).Value
        End If
        If Cells(outputRow, 12).Value > maxTotalVolume Then
            maxTotalVolume = Cells(outputRow, 12).Value
            maxTotalVolumeTicker = Cells(outputRow, 9).Value
        End If
    Next row
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = maxPercentageIncreaseTicker
    Cells(2, 17).Value = maxPercentageIncrease
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = maxPercentageDecreaseTicker
    Cells(3, 17).Value = maxPercentageDecrease
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = maxTotalVolumeTicker
    Cells(4, 17).Value = maxTotalVolume
    
    
    
End Sub
