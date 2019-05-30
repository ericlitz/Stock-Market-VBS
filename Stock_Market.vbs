Sub StockMarket()
    Dim stock_name As String
    
    Dim total_volume As Long
    total_volume = 0
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    For i = 2 To 70926
    
    stock_name = Cells(i, 1).Value
        
    total_volume = Cells(i, 7).Value
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Range("I" & summary_table_row).Value = stock_name
        
            Range("J" & summary_table_row).Value = total_volume
        
            summary_table_row = summary_table_row + 1
        
            total_volume = 0
        
        Else
        
            total_volume = total_volume + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub