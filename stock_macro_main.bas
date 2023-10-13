Attribute VB_Name = "Module1"
Sub macro_run()

Dim yearly_change As Double

For Each ws In ThisWorkbook.Worksheets

    'Adding labels to first section
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
      'Adding labels to second section
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
  
  
  
  
  
        i = 2    'summary row
        ticker = ws.Cells(2, 1)
        year_start_price = ws.Cells(2, 3)
        volume = ws.Cells(2, 7)
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For r = 3 To last_row

    
                If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
             
                year_end_price = ws.Cells(r, 6)
                volume = volume + ws.Cells(r, 7)
                yearly_change = (year_end_price - year_start_price)
                percentage_change = ((year_end_price - year_start_price) / year_start_price)

                
                'paste values into new table
                
                ws.Cells(i, 10).Value = ticker
                ws.Cells(i, 11).Value = yearly_change
                ws.Cells(i, 12).Value = percentage_change
                


                'ws.Cells(i, 12).Style = "Percent"
                ws.Cells(i, 13).Value = volume
                
                'Format Values
                
                    If ws.Cells(i, 11).Value >= 0 Then
                       ws.Cells(i, 11).Interior.ColorIndex = 4
                    
                        ElseIf ws.Cells(i, 11).Value < 0 Then
                        ws.Cells(i, 11).Interior.ColorIndex = 3
                    End If

                 'reset values
                 
                i = i + 1
                ticker = ws.Cells(r + 1, 1)
                year_start_price = ws.Cells(r + 1, 3)
                volume = 0
                
                Else
        
                volume = volume + ws.Cells(r, 7)
             
             End If

        Next r
        
            ws.Range("L2:L" & last_row).NumberFormat = "0.00%"
        
            top_percent_change = ws.Cells(2, 12)
            low_percent_change = ws.Cells(2, 12)
            top_ticker = ws.Cells(2, 10)
            low_ticker = ws.Cells(2, 10)
            most_volume_ticker = ws.Cells(2, 10)
            most_volume = ws.Cells(2, 13)
            
        
                For v = 3 To last_row
        
                     If ws.Cells(v, 12) > top_percent_change Then
                    top_percent_change = ws.Cells(v, 12)
                    top_ticker = ws.Cells(v, 10)
        
                    Else
                    End If
                    
                    
                             If ws.Cells(v, 12) < low_percent_change Then
                            low_percent_change = ws.Cells(v, 12)
                            low_ticker = ws.Cells(v, 10)
        
                            Else
                            End If
                    
                                    If ws.Cells(v, 13) > most_volume Then
                                        most_volume = ws.Cells(v, 13)
                                        most_volume_ticker = ws.Cells(v, 10)
        
                                    Else
                                    End If
                    
                Next v
        
                    ws.Cells(2, 17).Value = top_ticker
                    ws.Cells(3, 17).Value = low_ticker
                    ws.Cells(4, 17).Value = most_volume_ticker
        
                    ws.Cells(2, 18).Value = top_percent_change
                    ws.Cells(3, 18).Value = low_percent_change
                    ws.Cells(4, 18).Value = most_volume
                    
                    ws.Range("R2:R3").NumberFormat = "0.00%"
        
    Next ws
End Sub
