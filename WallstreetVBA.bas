Attribute VB_Name = "Module2"
Sub wallstreet()


For Each ws In Worksheets
ws.Activate

    
    'count tracks rows for columns i - l to specify which row to fill in during each loop
    Dim count As Integer
    count = 1
    'opening grabs the opening stock price year beginning
    Dim opening As Double
    opening = Cells(2, 3).Value
    'closing grabs the closing stock price year ending
    Dim closing As Double
    closing = 0
    'volume adds the total stock volume
    Dim volume As Variant
    volume = 0
    
  'finds and stores last row for use
   Dim LR As Variant
    LR = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    Dim i As Variant
    Dim j As Variant
    
    
    'labels columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
     
    

    'formats cells that need to be percents
    Range("K:K").NumberFormat = "0.00%"
    Range("Q2", "Q3").NumberFormat = "0.00%"

    
    For i = 2 To LR
    'fills in information in rows i through l when the value in column a changes
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            count = count + 1
            ws.Cells(count, 9) = ws.Cells(i, 1).Value
        
            volume = volume + ws.Cells(i, 7).Value
            ws.Cells(count, 12).Value = volume
            closing = ws.Cells(i, 6).Value
            ws.Cells(count, 10).Value = closing - opening
            
            If opening = 0 Then
                ws.Cells(count, 11).Value = 0
            Else
                ws.Cells(count, 11).Value = ws.Cells(count, 10) / opening
            End If
            
            ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
            opening = ws.Cells(i + 1, 3).Value
            closing = 0
            volume = 0
            
        Else
            volume = volume + ws.Cells(i, 7).Value
    End If
    
    Next i
    
    LR = Cells(Rows.count, 11).End(xlUp).Row
    
    For j = 2 To LR
    'conditional formatting colors for column j
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 3
        End If
        Next j
 

    'fills in greatest percent change
    For k = 2 To LR
    If Cells(k, 11).Value <> 0 Then
    If Cells(k, 11) > Cells(k + 1, 11) And Cells(k, 11).Value > Range("Q2").Value Then
    Range("Q2").Value = Cells(k, 11).Value
    Range("P2").Value = Cells(k, 9).Value
    End If
    End If
    Next k
    
    'fills in greatest percent decrease
    For r = 2 To LR
    If Cells(r, 11) < Cells(r + 1, 11) And Cells(r, 11).Value < Range("Q3").Value Then
    Range("Q3").Value = Cells(r, 11).Value
    Range("P3").Value = Cells(r, 9).Value
    End If
    Next r

'fills in greatest total volume
  For s = 2 To LR
    If Cells(s, 12).Value > Cells(s + 1, 12).Value And Cells(s, 12).Value > Range("Q4").Value Then
    Range("Q4").Value = Cells(s, 12).Value
    Range("P3").Value = Cells(s, 9).Value
    End If
    Next s
    
    
        
Next ws

End Sub

