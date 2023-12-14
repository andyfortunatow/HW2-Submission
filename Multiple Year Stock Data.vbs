Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    'Setting up the stage

        'Set Variable for the Ticker Symbol'
        Dim TickerSymbol As String
    
       'Creating a count for the Total Stock Volume as Total Stock Volume is the Sum of all cells in Column G within a Ticker Symbol Range'
        Dim StockVolume As Double
        StockVolume = 0

        'Keep track of the location for each ticker name in the summary table'
        Dim summaryrow As Integer
        summaryrow = 2
        
        'Note to self - The Yearly Change value is determined by Opening Price - Closing Price'
        'Note to self - The Percentage change is determined by (Closing Price - Opening Price)/Opening Price)*100'
        
        Dim Price1 As Double
        
        'Set Variable for 'AAB' Ticker to reference in retrieving'
        Price1 = Cells(2, 3).Value
        
        Dim Price2 As Double
        
        Dim Change As Double
        
        Dim Percentage_Change As Double

        'Label Table headers
        Cells(1, 14).Value = "Ticker"
        Cells(1, 15).Value = "Yearly Change"
        Cells(1, 16).Value = "Percent Change"
        Cells(1, 17).Value = "Total Stock Volume"

        'Counting the total number of rows in the Ticker Column'
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row

        'Looping through Ticker Column'
        For i = 2 To RowCount

        'Searching for the last Ticker symbol in the last cell before the symbol changes to a different symbol'
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the Ticker Symbol'
        TickerSymbol = Cells(i, 1).Value

        'Adding the total volume of that stock to the stock count'
        StockVolume = StockVolume + Cells(i, 7).Value

        'Inputting the symbol into the created table'
        Range("N" & summaryrow).Value = TickerSymbol

        'Inputting the volume into the created table'
        Range("Q" & summaryrow).Value = StockVolume

        'Setting the Closing Price value as the one five columns from the Ticker'
        Price2 = Cells(i, 6).Value

        'Calculate the yearly change
        Change = (Price2 - Price1)
              
        'Inputting yearly change for each ticker into the created table'
        Range("O" & summaryrow).Value = Change

        'Check for the non-divisibilty condition when calculating the percent change
        If (Price1 = 0) Then

        Percentage_Change = 0

        Else
                    
        Percentage_Change = Change / Price1
                
        End If

        'Inputting the yearly change for each ticker in the created table' - 'Sourced from Cited Example in README'
        Range("P" & summaryrow).Value = Percentage_Change
        Range("P" & summaryrow).NumberFormat = "0.00%"
   
        'Resettting Row Counter and Adding one to the Ticker Row' - 'Sourced from Cited Example in README'
        summaryrow = summaryrow + 1

        'Reset volume of trade to zero to ensure it does not add all ticker symbol's together- 'Sourced from Cited Example in README'
        StockVolume = 0

        'Reset the opening price to find new ticker opening price' - 'Sourced from Cited Example in README'
        Price1 = Cells(i + 1, 3)
            
        Else
              
        'Add the volume of trade - 'Sourced from Cited Example in README'
        StockVolume = StockVolume + Cells(i, 7).Value

            
        End If
        
        Next i

        'Determing the last row of the created to table to enter for conditional formatting'

        LastRow = Cells(Rows.Count, 15).End(xlUp).Row
    
        'Conditional formatting table'
    
        For i = 2 To LastRow
        If Cells(i, 15).Value > 0 Then
        
                Cells(i, 15).Interior.ColorIndex = 10
                
            Else
            
                Cells(i, 15).Interior.ColorIndex = 3
                
            End If
            
    Next i

End Sub

