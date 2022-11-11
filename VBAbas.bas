Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim wsname As String
        Dim i As Long
        Dim bob As Long
        Dim tickercnt As Long
        Dim lastrow As Long
        Dim lastrow1 As Long
        Dim changevar As Double
        
        wsname = ws.Name

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'ws.Cells(2, 15).Value = "Greatest % Increase"
        'ws.Cells(3, 15).Value = "Greatest % Decrease"
        'ws.Cells(4, 15).Value = "Greatest Total Volume"
        

        tickercnt = 2
        
        bob = 2
        

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To lastrow
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(tickercnt, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(tickercnt, 10).Value = ws.Cells(i, 6).Value - ws.Cells(bob, 3).Value
                
                    If ws.Cells(tickercnt, 10).Value < 0 Then
                
                    ws.Cells(tickercnt, 10).Interior.ColorIndex = 3
                
                    Else
                

                    ws.Cells(tickercnt, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(bob, 3).Value <> 0 Then
                    changevar = ((ws.Cells(i, 6).Value - ws.Cells(bob, 3).Value) / ws.Cells(bob, 3).Value)
                    
                    ws.Cells(tickercnt, 11).Value = Format(changevar, "Percent")
                    
                    Else
                    
                    ws.Cells(tickercnt, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(tickercnt, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(bob, 7), ws.Cells(i, 7)))
                

                tickercnt = tickercnt + 1
                
                bob = i + 1
                
                End If
            
            Next i

        
        
            
    Next ws
        
End Sub
