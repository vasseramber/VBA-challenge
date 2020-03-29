Sub Multiple_year_stock_data():
    
    
    Dim ws As Worksheet
    Dim total As Double
    Dim j As Integer
    
    For Each ws In Worksheets
    
        
        total = 0
        j = 0
           
         
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        'RowCount = 797711
        
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        For i = 2 To RowCount
        
           
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            
                
                ws.Range("J" & 2 + j).Value = total
               
                
                total = 0
               
                
                j = j + 1
            
                
            Else
                total = total + Cells(i, 7).Value
            End If
        
          Next i
    
    Next ws

End Sub
