Attribute VB_Name = "Module1"

Sub StockData()

        Dim WS As Worksheet
        Dim Quarterlychange As Double
        Dim Percentchange As Double
        Dim Totalstockvolume As Long
        Dim Ticker As String
        Dim Openprice As Double
        Dim Closeprice As Double
        Dim i As Long
        Dim lastrow As Long
        Dim Greatestincrease As Double
        Dim Greatestdecrease As Double
        Dim Greatesttotalva As Long
        Dim increaseRow As Long
        Dim decreaseRow As Long
        Dim volumeRow As Long
        Dim r As Range
        Dim Totalstockvalue As Long
     
        
        
        
        
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.
'loop through each WS
        For Each WS In Worksheets
        
       
    
'Create headers
      
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Quarterly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volume"
    WS.Range("O2").Value = "Greatest % Increase"
    WS.Range("O3").Value = "Greatest % Decrease"
    WS.Range("O4").Value = "Greatest Total Value"
    WS.Range("P1").Value = "Ticker"
    WS.Range("Q1").Value = "Value"
        
    
    WS.Columns("I:M").EntireColumn.AutoFit
    WS.Columns("O:Q").EntireColumn.AutoFit
            
'Ticker column

    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    outputrow = 2
    For i = 2 To lastrow
    
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
            Ticker = WS.Cells(i, 1).Value
            WS.Cells(outputrow, 9).Value = Ticker
        
'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
        If WS.Cells(i, 3).Value <> 0 Then
            Quarterlychange = (WS.Cells(i, 6).Value - WS.Cells(i, 3).Value)
            WS.Cells(outputrow, 10).Value = Quarterlychange
            Quarterlychange = 0
            
'The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
            
            percentagechange = (WS.Cells(i, 6).Value - WS.Cells(i, 3).Value) / WS.Cells(i, 3).Value
            WS.Cells(outputrow, 11).Value = percentagechange
            WS.Columns("K:K").NumberFormat = "0.00%"
            percentagechange = 0
          
'The total stock volume of the stock.
        
           Totalstockvolume = Totalstockvolume + WS.Cells(i, 7).Value
           WS.Cells(outputrow, 12).Value = Totalstockvolume
     outputrow = outputrow + 1
     
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        Greatestincrease = -1
        Greatestdecrease = 1
        Greatesttotalvalue = 0
        

         If WS.Cells(i, 11).Value > Greatestincrease Then
            Greatestincrease = WS.Cells(i, 11).Value
            increaseRow = n
        End If

        If WS.Cells(i, 11).Value < Greatestdecrease Then
            Greatestdecrease = WS.Cells(i, 11).Value
            decreaseRow = n
        End If

        If WS.Cells(i, 12).Value > Greatestvolume Then
             Greatestvolume = WS.Cells(i, 12).Value
            volumeRow = n
        End If
           
        
    WS.Cells(2, 16).Value = WS.Cells(increaseRow, 1).Value  ' Greatest % Increase Ticker
    WS.Cells(2, 17).Value = Greatestincrease ' Greatest % Increase Value

    WS.Cells(3, 16).Value = WS.Cells(decreaseRow, 1).Value ' Greatest % Decrease Ticker
    WS.Cells(3, 17).Value = Greatestdecrease ' Greatest % Decrease Value

    WS.Cells(4, 16).Value = WS.Cells(volumeRow, 1).Value ' Greatest Total Volume Ticker
    WS.Cells(4, 17).Value = Greatestvolume ' Greatest Total Volume Value


'Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
       
    Set r = WS.Range("J2:J" & lastrow)
    
     With r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = vbGreen ' Green for positive changes
    End With

    With r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = vbRed ' Red for negative changes
    End With
    
          End If
          
                     
           End If
            
       
                      
        
    Next i
 
  

        Next
 
End Sub



