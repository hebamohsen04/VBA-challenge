Sub TickerAnalysis()

  ' Create a variable to hold the TickerAnalysis .
  Dim i As Double
  Dim j As Double
  Dim lastrow As Long
  Dim yearOpeningValue As Double
  Dim yearClosingValue As Double
  Dim totalVolume As Double
  
  Dim greatestVolumeTicker As String
  Dim greatestVolume As Double
  
  Dim greatestIncreaseTicker As String
  Dim greatestIncrease As Double
  
  Dim greatestDecreaseTicker As String
  Dim greatestDecrease As Double
  
  j = 2
  
  ' counts the number of rows
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
 Range("L1").Value = "Total Stock Volume"
 
 yearOpeningValue = Cells(2, 3).Value
 totalVolume = 0
 greatestVolume = 0
 greatestIncrease = 0
 greatestDecrease = 0
 
  For i = 2 To lastrow
  
  totalVolume = totalVolume + Cells(i, 7).Value
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then ' We Finished the previous ticker
        yearClosingValue = Cells(i, 6).Value
        
        Cells(j, 9).Value = Cells(i, 1).Value
        Cells(j, 10).Value = yearClosingValue - yearOpeningValue
        
        If yearClosingValue > 0 Then
            Cells(j, 11).Value = (yearClosingValue - yearOpeningValue) / yearClosingValue
            If greatestIncrease <= (yearClosingValue - yearOpeningValue) / yearClosingValue Then
                greatestIncrease = (yearClosingValue - yearOpeningValue) / yearClosingValue
                greatestIncreaseTicker = Cells(i, 1).Value
            End If
            
            If greatestDecrease >= (yearClosingValue - yearOpeningValue) / yearClosingValue Then
                greatestDecrease = (yearClosingValue - yearOpeningValue) / yearClosingValue
                greatestDecreaseTicker = Cells(i, 1).Value
            End If
            
        
        Else
            Cells(j, 11).Value = 0
        End If
        
        Cells(j, 11).NumberFormat = "0.00%"
        Cells(j, 12).Value = totalVolume
        
        If totalVolume >= greatestVolume Then
           greatestVolume = totalVolume
           greatestVolumeTicker = Cells(i, 1).Value
        End If
        
            
        
        
        If (yearClosingValue - yearOpeningValue) >= 0 Then
            Cells(j, 10).Interior.Color = RGB(0, 255, 119)
        Else
            Cells(j, 10).Interior.Color = RGB(255, 0, 0)
        End If
        
        j = j + 1
         
         yearOpeningValue = Cells(i + 1, 3).Value
         totalVolume = 0
    
    End If
  Next i
  
  Range("N2").Value = "Greatest % Increase"
  Range("N3").Value = "Greatest % Decrease"
  Range("N4").Value = "Greatest Total Volume"
  
  Range("O4").Value = greatestVolumeTicker
  Range("P4").Value = greatestVolume
  
  Range("O2").Value = greatestIncreaseTicker
  Range("P2").Value = greatestIncrease
  Range("P2").NumberFormat = "0.00%"
  
  Range("O3").Value = greatestDecreaseTicker
  Range("P3").Value = greatestDecrease
   
   Range("P3").NumberFormat = "0.00%"


End Sub
