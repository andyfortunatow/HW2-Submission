Attribute VB_Name = "Module2"
Sub Bonus():

 'Label Table headers
        
        Cells(1, 21).Value = "TickerSymbol"
        Cells(1, 22).Value = "Value"
        Cells(2, 20).Value = "Greatest % Increase"
        Cells(3, 20).Value = "Greatest % Decrease"
        Cells(4, 20).Value = "Greatest Total Volume"

'Determing the last row of the created to table to enter in loop

        LastRow = Cells(Rows.Count, 15).End(xlUp).Row
        
        MsgBox (LastRow)
        
        LastRow = 3001

'Setting Variable for Greatest_Increase'

        Dim Greatest_Increase As Double
 
        Greatest_Increase = WorksheetFunction.Max(Range("P2:P3001"))

        MsgBox (Greatest_Increase)

'Setting Variable for Greatest_Decrease'

        Dim Greatest_Decrease As Double
 
        Greatest_Decrease = WorksheetFunction.Min(Range("P2:P3001"))

        MsgBox (Greatest_Decrease)
        

'Setting Variable for Greatest_Decrease'

        Dim Greatest_StockVolume As Double
 
        Greatest_StockVolume = WorksheetFunction.Max(Range("Q2:Q3001"))

        MsgBox (Greatest_StockVolume)

'Creating Loop to Input Greatest Increase, Greatest Decrease and Total Stock Volume and Relevant Ticker Symbol'

        For i = 2 To LastRow
    
        If Cells(i, 16).Value = Greatest_Increase Then
        
        Range("V2").Value = Cells(i, 16)
        
        End If
        
        If Cells(i, 16).Value = Greatest_Increase Then
        
        Range("U2").Value = Cells(i, 14)
        
        End If
        
        If Cells(i, 16).Value = Greatest_Decrease Then
        
        Range("V3").Value = Cells(i, 16)
        
        End If
        
        If Cells(i, 16).Value = Greatest_Decrease Then
        
        Range("U3").Value = Cells(i, 14)
        
        End If
        
        If Cells(i, 17).Value = Greatest_StockVolume Then
        
        Range("V4").Value = Cells(i, 17)
        
        End If
        
        If Cells(i, 17).Value = Greatest_StockVolume Then
        
        Range("U4").Value = Cells(i, 14)
        
        End If
        
        Next i
        
'Formatting V2 and V3 as Percentage'
        
        Range("V2:V3").NumberFormat = "0.00%"

End Sub

