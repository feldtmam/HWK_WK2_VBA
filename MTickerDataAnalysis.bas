Attribute VB_Name = "MTickerDataAnalysis"
Sub StockTicker()

Dim ws As Worksheet
Dim NumRows, WriteRow, NumRowsAggregated As Integer
Dim TickerSymbol, HighestPercIncreaseTicker, HighestPercDecreaseTicker, GreatestTotalVolumeTicker As String
Dim StartPrice, TotalChange, PercentageChange, EndPrice, HighestPercIncreaseValue, HighestPercDecreaseValue As Double
Dim TotalStockVolume, GreatestTotalVolumeValue As Double

'Switch off screen updating and auto calculate to improve performance
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Loop through each worksheet - name of the sheet is the year value
For Each ws In ActiveWorkbook.Worksheets
   ws.Activate
   With ws
       'Write the headers to the relevant cells
       .Cells(1, 10).Value = "Ticker"
       .Cells(1, 11).Value = "Yearly Change"
       .Cells(1, 12).Value = "Percent Change"
       .Cells(1, 13).Value = "Total Stock Volume"
       .Cells(1, 17).Value = "Ticker"
       .Cells(1, 18).Value = "Value"
       .Cells(2, 16).Value = "Greatest % Increase"
       .Cells(3, 16).Value = "Greatest % Decrease"
       .Cells(4, 16).Value = "Greatest Total Volume"

       'Identify the number of rows to walk through
       NumRows = Cells(Rows.Count, 1).End(xlUp).Row
       'Data has already been sorted, so can use the first ticker symbol to start with
       TickerSymbol = .Cells(2, 1).Value
       StartPrice = .Cells(2, 3).Value
       WriteRow = 2
       TotalStockVolume = 0
       'Start at the 2nd row and assign the ticker symbol, Opening Price and Percentage
       For i = 2 To NumRows
           'Loop through each row while the ticker symbol in the row is the same as the ticker symbol variable
           'When ticker symbol is different then write the total ticker name and stock volume to the relevant row, and do and write calculations for total change and percentage change
           While Cells(i, 1).Value <> TickerSymbol 'If the symbol is different, then change the values, otherwise add up the values
               'Write the values to the relevan row
               .Cells(WriteRow, 10).Value = TickerSymbol
               .Cells(WriteRow, 13).Value = TotalStockVolume
               'Calculate the difference between the start price and end price for the ticker
               TotalChange = EndPrice - StartPrice

               'Assign 0 to the percentage change when the denominator is 0, otherwise error is generated
               If StartPrice = 0 Then
                   PercentageChange = 0
               Else
                   'Calculate the percentage change between the start and end values
                   PercentageChange = (TotalChange / StartPrice) * 100
               End If
               'Write the calculated values to the relevant row
               .Cells(WriteRow, 11).Value = TotalChange
               .Cells(WriteRow, 12).Value = PercentageChange
               'Increase the row value for the next set of aggregated values and start stock volume at 0
               WriteRow = WriteRow + 1
               TotalStockVolume = 0
               'Assign the next ticker symbo and startvalue
               TickerSymbol = .Cells(i, 1).Value
               StartPrice = .Cells(i, 3).Value

           Wend
           'While the ticker symbol is still the same, add up the total stock volume and assign the endprice as the current end price
           'Counter to determine the total stock volume
           TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
           EndPrice = .Cells(i, 6).Value
       Next i

       'Determine the number of rows written on the aggregated data
       NumRowsAggregated = Cells(Rows.Count, 10).End(xlUp).Row
       'Assign 0 to the variables to start with
       HighestPercIncreaseValue = 0
      HighestPercDecreaseValue = 0
       GreatestTotalVolumeValue = 0
       'Loop through the aggregated data and identify the stock with the greates % increase decrease and total volume
       For k = 2 To NumRowsAggregated
           'Find and write the Highest Percentage increase values
           If HighestPercIncreaseValue < .Cells(k, 12).Value Then
               HighestPercIncreaseValue = .Cells(k, 12).Value
               HighestPercIncreaseTicker = .Cells(k, 10).Value
           End If
           'Find and write the Highest Percentage decrease values
           If .Cells(k, 12).Value < 0 And HighestPercDecreaseValue > .Cells(k, 12).Value Then
               HighestPercDecreaseValue = .Cells(k, 12).Value
               HighestPercDecreaseTicker = .Cells(k, 10).Value
           End If
           'Find and write the Highest volume increase values
           If GreatestTotalVolumeValue < .Cells(k, 13).Value Then
               GreatestTotalVolumeValue = .Cells(k, 13).Value
               GreatestTotalVolumeTicker = .Cells(k, 10).Value
           End If
           'Assign the colors to the difference between start and end price cells
           If .Cells(k, 11).Value < 0 Then
                   .Cells(k, 11).Interior.ColorIndex = 3
               ElseIf .Cells(k, 11).Value = 0 Then 'To ensure we're not dividing by 0 Ive assigned 0 when the start price was 0, so giving the 0 a different color
                   .Cells(k, 11).Interior.ColorIndex = 7
               Else
                   .Cells(k, 11).Interior.ColorIndex = 4
               End If
       Next k
       'Write the valuse from the variables above to the correct cells
       .Cells(2, 17).Value = HighestPercIncreaseTicker
       .Cells(2, 18).Value = HighestPercIncreaseValue
       .Cells(3, 17).Value = HighestPercDecreaseTicker
       .Cells(3, 18).Value = HighestPercDecreaseValue
       .Cells(4, 17).Value = GreatestTotalVolumeTicker
       .Cells(4, 18).Value = GreatestTotalVolumeValue

   End With
Next ws

MsgBox "Analysis Completed"

'Enable screen updating and auto calculation again
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
