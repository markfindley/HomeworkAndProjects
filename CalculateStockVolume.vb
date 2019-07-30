Sub CalculateStockVolume()

    Dim ColumnName1 As String
    Dim ColumnName2 As String
    Dim ColumnName3 As String
    Dim ColumnName4 As String
    Dim ColumnName5 As String
    Dim ColumnName6 As String
    
    Dim RowName1 As String
    Dim RowName2 As String
    Dim RowName3 As String

    ColumnName1 = "Ticker"
    ColumnName2 = "Yearly Change"
    ColumnName3 = "Percent Change"
    ColumnName4 = "Total Stock Volume"
    ColumnName5 = "Ticker"
    ColumnName6 = "Value"
    
    RowName1 = "Greatest % Increase"
    RowName2 = "Greatest % Decrease"
    RowName3 = "Greatest Total Volume"


    For Each ws In Worksheets

        ws.Cells(1, 9).Value = ColumnName1
        ws.Cells(1, 10).Value = ColumnName2
        ws.Cells(1, 11).Value = ColumnName3
        ws.Cells(1, 12).Value = ColumnName4
        ws.Cells(1, 16).Value = ColumnName5
        ws.Cells(1, 17).Value = ColumnName6

        Dim WorksheetName As String
        Dim LastRow As Long
        Dim TotalSalesVolume As Double
        Dim StockTicker As String
        Dim StockTickerCount As Integer
        Dim OpeningYearPrice As Currency
        Dim ClosingYearPrice As Currency
        Dim YearlyPriceChange As Currency
        Dim YearlyPricePercentageChange As Double
        Dim GreatestPercentageIncrease As Double
        Dim GreatestPercentageDecrease As Double
        Dim GreatestTotalVolume As Double
        Dim GreatestPercentageIncreaseTicker As String
        Dim GreatestPercentageDecreaseTicker As String
        Dim GreatestTotalVolumeTicker As String

        WorksheetName = ws.Name
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' First clear the worksheets
        ws.Range("I2:I" & LastRow).Value = ""
        ws.Range("J2:J" & LastRow).Value = ""
        ws.Range("K2:K" & LastRow).Value = ""
        ws.Range("L2:L" & LastRow).Value = ""
        ws.Range("O2:O" & LastRow).Value = ""
        ws.Range("P2:P" & LastRow).Value = ""
        ws.Range("Q2:Q" & LastRow).Value = ""
        ws.Range("I2:I" & LastRow).ClearFormats
        ws.Range("J2:J" & LastRow).ClearFormats
        ws.Range("K2:K" & LastRow).ClearFormats
        ws.Range("L2:L" & LastRow).ClearFormats
        ws.Range("O2:O" & LastRow).ClearFormats
        ws.Range("P2:P" & LastRow).ClearFormats
        ws.Range("Q2:Q" & LastRow).ClearFormats
        
        StockTicker = ws.Cells(2, 1).Value
        OpeningYearPrice = ws.Cells(2, 3).Value
        StockTickerCount = 2
        GreatestPercentageIncrease = 0
        GreatestPercentageDecrease = 0
        GreatestTotalVolume = 0

        For I = 2 To LastRow

            If ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value Then ' Need to change everytime a new ticker symbol appears

                TotalSalesVolume = TotalSalesVolume + ws.Cells(I + 1, 7).Value

            Else ' Now put the metrics on each sheet

                ClosingYearPrice = ws.Cells(I, 6).Value
                YearlyPriceChange = ClosingYearPrice - OpeningYearPrice
                
                If OpeningYearPrice <> 0 Then
                    YearlyPricePercentageChange = (ClosingYearPrice / OpeningYearPrice) - 1
                Else
                    YearlyPricePercentageChange = 0
                End If
                
                ' positive change in green and negative change in red
                ws.Range("I" & StockTickerCount).Value = StockTicker
                ws.Range("J" & StockTickerCount).Value = YearlyPriceChange
                ws.Range("K" & StockTickerCount).Value = YearlyPricePercentageChange
                ws.Range("L" & StockTickerCount).Value = TotalSalesVolume
                
                ws.Range("K" & StockTickerCount).NumberFormat = "0.00%"
                
                If YearlyPriceChange > 0 Then
                    ws.Range("J" & StockTickerCount).Interior.ColorIndex = 4
                ElseIf YearlyPriceChange < 0 Then
                    ws.Range("J" & StockTickerCount).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & StockTickerCount).Interior.ColorIndex = 0
                End If
                
                If YearlyPricePercentageChange > GreatestPercentageIncrease Then
                    GreatestPercentageIncrease = YearlyPricePercentageChange
                    GreatestPercentageIncreaseTicker = StockTicker
                End If
                
                If YearlyPricePercentageChange < GreatestPercentageDecrease Then
                    GreatestPercentageDecrease = YearlyPricePercentageChange
                    GreatestPercentageDecreaseTicker = StockTicker
                End If
                
                If TotalSalesVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalSalesVolume
                    GreatestTotalVolumeTicker = StockTicker
                End If
                
                TotalSalesVolume = 0
                StockTicker = ws.Cells(I + 1, 1).Value
                OpeningYearPrice = ws.Cells(I + 1, 3).Value
                StockTickerCount = StockTickerCount + 1
                
            End If
             
        Next I
        
        ws.Range("O2").Value = RowName1
        ws.Range("O3").Value = RowName2
        ws.Range("O4").Value = RowName3
        ws.Range("P2").Value = GreatestPercentageIncreaseTicker
        ws.Range("P3").Value = GreatestPercentageDecreaseTicker
        ws.Range("P4").Value = GreatestTotalVolumeTicker
        ws.Range("Q2").Value = GreatestPercentageIncrease
        ws.Range("Q3").Value = GreatestPercentageDecrease
        ws.Range("Q4").Value = GreatestTotalVolume
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("I1:I" & LastRow).Columns.AutoFit
        ws.Range("J1:J" & LastRow).Columns.AutoFit
        ws.Range("K1:K" & LastRow).Columns.AutoFit
        ws.Range("L1:L" & LastRow).Columns.AutoFit
        
        ws.Range("O1:O" & LastRow).Columns.AutoFit
        ws.Range("P1:P" & LastRow).Columns.AutoFit
        ws.Range("Q1:Q" & LastRow).Columns.AutoFit

    Next ws
    
End Sub
