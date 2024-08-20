Sub StockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim volume As Double
    Dim outputRow As Long
    Dim i As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVol As Double

    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables for each worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputRow = 2
        volume = 0
        
        ' Set up headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize variables for greatest values
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVol = 0

        ' Loop through all rows
        For i = 2 To lastRow
            If i = 2 Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' New ticker found, set opening price
                Ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                volume = 0
            End If
            
            ' Accumulate volume
            volume = volume + ws.Cells(i, 7).Value
            
            ' Check if it's the last row of the current ticker
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set closing price and calculate changes
                closingPrice = ws.Cells(i, 6).Value
                quarterlyChange = closingPrice - openingPrice
                
                ' Calculate percentage change, avoiding division by zero
                If openingPrice <> 0 Then
                    percentageChange = quarterlyChange / openingPrice
                Else
                    percentageChange = 0
                End If
                
                ' Output results
                ws.Cells(outputRow, 9).Value = Ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 12).Value = volume
                
                ' Format percentage
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                ' Color formatting
                If quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                Else
                    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                End If
                
                ' Check for greatest values
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    ws.Cells(2, 16).Value = Ticker
                    ws.Cells(2, 17).Value = greatestIncrease
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                End If
                
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    ws.Cells(3, 16).Value = Ticker
                    ws.Cells(3, 17).Value = greatestDecrease
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                End If
                
                If volume > greatestVol Then
                    greatestVol = volume
                    ws.Cells(4, 16).Value = Ticker
                    ws.Cells(4, 17).Value = greatestVol
                End If
                
                ' Move to next output row
                outputRow = outputRow + 1
            End If
        Next i
        
        ' Set labels for greatest values
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Autofit columns
        ws.Columns("I:Q").AutoFit
    Next ws
End Sub
