Attribute VB_Name = "Module1"
Sub Multiple_year_stock()
    Dim lastRow As Long
    Dim i As Long
    Dim percentChange As Double
    Dim maxPercentIncrease As Double
    Dim minPercentDecrease As Double
    Dim maxVolume As Double
    Dim maxPercentIncreaseChange As String
    Dim minPercentDecreaseChange As String
    Dim maxVolumeChange As String
    Dim maxPercentIncreaseTicker As String
    Dim minPercentDecreaseTicker As String
    Dim tinkerValue As String
    Dim maxPercentage As Double
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
  
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Quarterly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    

    For i = 2 To lastRow
     
        Cells(i, "I").Value = Cells(i, "A").Value
        
    
        Cells(i, "J").Value = Cells(i, "C").Value - Cells(i, "F").Value
        
    
        If Cells(i, "C").Value = 0 Then
            Cells(i, "K").Value = "N/A"
        Else
            percentChange = ((Cells(i, "F").Value - Cells(i, "C").Value) / Cells(i, "C").Value) * 100
            Cells(i, "K").Value = percentChange & "%"
      
            If Cells(i, "J").Value > 0 Then
                Cells(i, "J").Interior.Color = RGB(0, 255, 0)
            ElseIf Cells(i, "J").Value < 0 Then
                Cells(i, "J").Interior.Color = RGB(255, 0, 0)
            End If
        End If
        
 
        Cells(i, "L").Value = Cells(i, "G").Value * (1 + percentChange / 100)
        
       
        If Abs(Cells(i, "K").Value) > maxPercentage Then
            maxPercentage = Abs(Cells(i, "K").Value)
            tinkerValue = Cells(i, "A").Value
        End If
      
        If Cells(i, "K").Value > maxPercentIncrease Then
            maxPercentIncrease = Cells(i, "K").Value
            maxPercentIncreaseTicker = Cells(i, "A").Value
        ElseIf Cells(i, "K").Value < minPercentDecrease Then
            minPercentDecrease = Cells(i, "K").Value
            minPercentDecreaseTicker = Cells(i, "A").Value
        End If
       
        If Cells(i, "L").Value > maxVolume Then
            maxVolume = Cells(i, "L").Value
            maxVolumeTicker = Cells(i, "A").Value
        End If
    Next i
   
    Range("P2").Value = tinkerValue
    
    Range("Q1").Value = "Value"
    Range("Q2").Value = Format(maxPercentIncrease, "0.00%")
    Range("Q3").Value = Format(minPercentDecrease, "0.00%")
    Range("Q4").Value = maxVolume
    Range("O2").Value = "Greatest% Increase"
    Range("O3").Value = "Greatest% Decrease"
    Range("O4").Value = "Greatest% Volume"
    Range("P3").Value = minPercentDecreaseTicker
    Range("P4").Value = maxVolumeTicker
    Next ws
End Sub
