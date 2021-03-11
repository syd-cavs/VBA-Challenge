Sub VBA_Challenge2():

    Dim ws As Worksheet

    For Each ws In Worksheets
    
        Dim ticker As String
        
        Dim volume As Double
        volume = 0
        
        Dim table_row As Integer
        table_row = 2
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Bonus Code
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        'Bonus Code
        
        Dim i As Long
        
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        Dim open_change As Double
        open_change = ws.Cells(2, 3).Value
        Dim close_change As Double
        close_change = 0
        Dim yearly_change As Double
        yearly_change = 0
        
        Dim percent_change As Double
        percent_change = 0
        
        'Bonus code
        Dim max_increase As Double
        max_increase = 0
        Dim stock_name As String
        Dim stock_name1 As String
        Dim stock_name2 As String
        Dim min_increase As Double
        min_increase = 0
        Dim max_volume As Double
        max_volume = 0
        'Bonus Code
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                volume = volume + ws.Cells(i, 7).Value
                close_change = ws.Cells(i, 6).Value
                yearly_change = close_change - open_change
                
                If open_change = 0 Then
                    percent_change = 0
                Else
                    percent_change = Round((yearly_change / open_change) * 100, 2)
                    
                End If
                
                If yearly_change >= 0 Then
                    ws.Range("J" & table_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & table_row).Interior.ColorIndex = 3
                End If
                
                ws.Range("I" & table_row).Value = ticker
                ws.Range("L" & table_row).Value = volume
                ws.Range("J" & table_row).Value = yearly_change
                ws.Range("K" & table_row).Value = percent_change & "%"
                
                table_row = table_row + 1
                volume = 0
                yearly_change = 0
                percent_change = 0
                open_change = ws.Cells(i + 1, 3).Value
            Else
                volume = volume + ws.Cells(i, 7).Value
        
            End If
            
            'Bonus Code
            If ws.Cells(i, 11) > max_increase Then
                max_increase = ws.Cells(i, 11).Value
                stock_name = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11) < min_increase Then
                min_increase = ws.Cells(i, 11).Value
                stock_name1 = ws.Cells(i, 9).Value
            Else
            End If
            
            If ws.Cells(i, 12) > max_volume Then
                max_volume = ws.Cells(i, 12).Value
                stock_name2 = ws.Cells(i, 9).Value
            Else
            End If
        
            ws.Range("O2").Value = stock_name
            ws.Range("P2").Value = max_increase & "%"
            ws.Range("O3").Value = stock_name1
            ws.Range("P3").Value = min_increase & "%"
            ws.Range("O4").Value = stock_name2
            ws.Range("P4").Value = max_volume & "%"
            'Bonus Code
            
        Next i
                
    Next ws
    
    
End Sub
