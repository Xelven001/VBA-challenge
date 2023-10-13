Attribute VB_Name = "Module2"
Sub macro_reset()

    For Each ws In ThisWorkbook.Worksheets
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row


        ws.Range("J1:R" & last_row).Value = ""
        ws.Range("J1:R" & last_row).Interior.ColorIndex = xlNone
    Next ws
End Sub
