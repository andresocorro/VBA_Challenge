Attribute VB_Name = "Module1"
Sub VBA_Stocks()

    For Each ws In Worksheets

    '   SET HEADERS
    
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Stock Volume"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
    
    ' DIM VARIABLES
    
        Dim stock_ticker As String
        Dim year_open As Double
        Dim year_close As Double
        Dim year_change As Double
        Dim year_percentage As Double
        Dim total_vol As LongLong
        total_vol = 0
    
        Dim summary_row As Integer
        summary_row = 2
        Dim openprice As Long
        openprice = 2
       
    ' FIND LAST ROW
       
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
    ' FIND OUT TICKER SYMBOLS & LAST VOLUME IN I
    
        stock_ticker = ws.Cells(i, 1).Value
        total_vol = total_vol + ws.Cells(i, 7).Value
        
    ' CALCULATE YEARLY CHANGE
        year_open = ws.Cells(openprice, 3).Value
        year_close = ws.Cells(i, 6).Value
        year_change = year_close - year_open
        
        ' CALCULATE PERCENTAGE CHANGE
    
         If (year_open = 0 And year_change = 0) Or (year_open = 0 And year_change <> 0) Then
         year_open = 1
        
         year_percentage = (year_change / year_open)
        
         Else
        
         year_percentage = (year_change / year_open)
        
        End If
        
    ' PLACE RESULTS IN THEIR RESPECTIVE CELLS
        ws.Range("I" & summary_row).Value = stock_ticker
        ws.Range("J" & summary_row).Value = year_change
        ws.Range("K" & summary_row).Value = year_percentage
        ws.Range("L" & summary_row).Value = total_vol
        
         ' Format Percent change percentage
            ws.Range("K" & summary_row).Value = Format(year_percentage, "Percent")
         ' Conditional format Yearly change
            If year_change > 0 Then
            
            ws.Range("J" & summary_row).Interior.ColorIndex = 4
            
            ElseIf year_change < 0 Then
            
            ws.Range("J" & summary_row).Interior.ColorIndex = 3
            
            Else
            
            ws.Range("J" & summary_row).Interior.ColorIndex = 6
            
            End If
            
    ' RESTART COUNTERS TO FIND NEXT VALUE
        summary_row = summary_row + 1
        total_vol = 0
        openprice = i + 1
        
    ' ADD ALL THE VOLUMES WHILE i REMAINS EQUAL TO i+1
    
        Else
            total_vol = total_vol + ws.Cells(i, 7).Value
        
        End If
    
        Next i
    
 ' FIND BIGGEST % INCREASE AND % DECREASE
   
    ws.Cells(2, 17).Value = Format(Application.WorksheetFunction.Max(ws.Range("K:K")), "Percent")
    ws.Cells(3, 17).Value = Format(Application.WorksheetFunction.Min(ws.Range("K:K")), "Percent")
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
   
 
   lastrow_percentage = ws.Cells(Rows.Count, 11).End(xlUp).Row
   
 
   
    For i = 2 To lastrow_percentage
    
        If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
        ws.Range("P4").Value = ws.Cells(i, 9).Value
        
        End If
    
    Next i
   
        ws.Cells.EntireColumn.AutoFit
        
Next ws


End Sub

