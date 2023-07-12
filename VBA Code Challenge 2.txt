Sub stockdatasummary()
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
  ws.Activate

    
    Dim ticker_name As String

    
    Dim total_volume As Double
    total_volume = 0

    
    Dim ticker_row As Integer
    ticker_row = 2
  
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
       
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
  
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
    
       For i = 2 To lastrow
        
        If Cells(i - 1, 1) <> Cells(i, 1) Then
        Opening_Price = Cells(i, 3)
        
       
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        
        ticker_name = Cells(i, 1).Value

        
        total_volume = total_volume + Cells(i, 7).Value
      
        
        closing_price = Cells(i, 6).Value
      
        
        yearly_change = closing_price - Opening_Price
      
        
        percent_change = ((closing_price - Opening_Price) / Opening_Price)
        On Error Resume Next

        
        Range("I" & ticker_row).Value = ticker_name
      
        
        Range("J" & ticker_row).Value = yearly_change
      
        
        Range("K" & ticker_row).Value = percent_change
        Columns("K:K").NumberFormat = "0.00%"

        
        Range("L" & ticker_row).Value = total_volume
      
        
        ticker_row = ticker_row + 1
      
        
        total_volume = 0

        
        Else
        
        total_volume = total_volume + Cells(i, 7).Value

        End If
             
       Next i
        
        
       Dim greatest_increase As Double
       Dim greatest_decrease As Double
       greatest_increase = Cells(2, 11)
       greatest_decrease = Cells(2, 11)
       greatest_volume = Cells(2, 12)
       lastrow_summary = Cells(Rows.Count, 10).End(xlUp).Row
  
       For j = 2 To lastrow_summary
        
        If Cells(j, 10) >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
  
        ElseIf Cells(j, 10) < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
   
        End If
        
        
        If Cells(j, 11) > greatest_increase Then
        greatest_increase = Cells(j, 11)
        Cells(2, 17) = greatest_increase
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16) = Cells(j, 9)
   
        End If
   
        
        If Cells(j, 11) < greatest_decrease Then
        greatest_decrease = Cells(j, 11)
        Cells(3, 17) = greatest_decrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16) = Cells(j, 9)
  
        End If
        
        
        If Cells(j, 12) > greatest_volume Then
        greatest_volume = Cells(j, 12)
        Cells(4, 17) = greatest_volume
        Cells(4, 16) = Cells(j, 9)
   
        End If
   
       Next j
 
       
 Next

End Sub

