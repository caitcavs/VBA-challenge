Sub Analyze_Stocks()

Dim i As Long
Dim summaryrowcounter As Double
Dim TotalVolume As Double
Dim percentchange As Single
Dim openingprice As Variant
Dim yearlychange As Variant

TotalVolume = 0
yearlychange = 0
Start = 2
summaryrowcounter = 2
openingprice = Cells(2, 3)
    'Determine row count of the last row with data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through all rows of stock data
    For i = 2 To RowCount
        TotalVolume = TotalVolume + Cells(i, 7).Value
        'Assign closing price to column of cells
        closingprice = Cells(i, 6)
        'Yearly change will subtract opening price from closing price
        yearlychange = closingprice - openingprice
        'Assign changing variable of opening price
        openingprice = Cells(i + 1, 3)
        'Assign how to get percent change values
        percentchange = yearlychange / openingprice
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            
            'Assign cells to receive ticker counter values
            Cells(summaryrowcounter, 9) = Cells(i, 1)
            
            
            'Assign cells to receive Total Volume values
            Cells(summaryrowcounter, 12).Value = TotalVolume
            
            'Assign cells to receive Yearly Change values
            Cells(summaryrowcounter, 10).Value = yearlychange
            
            'Assign cells to receive Percent Change values
            Cells(summaryrowcounter, 11).Value = percentchange
                
            
            
            'Resets total volume back to zero for next unique ticker
            TotalVolume = 0
            
            'Allows loop through next unique ticker row
            summaryrowcounter = summaryrowcounter + 1

               
            
        End If

    Next
End Sub

