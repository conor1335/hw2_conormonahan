Attribute VB_Name = "Module1"
Sub medium()


' Create a script that will loop through all the stocks and take the following info.

' Yearly change from what the stock opened the year at to what the closing price was.
' The percent change from the what it opened the year at to what it closed.
' The total Volume of the stock
' Ticker Symbol
' You should also have conditional formatting that will highlight
' positive change in green and negative change in red


For Each ws In Worksheets
    Dim WorksheetName As String
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    lastrow2 = ws.Cells(Rows.Count, 1).End(xlUp).row
    WorksheetName = ws.Name
    
    Dim row As Long
    
    Dim curr_ticker_count As Long
    curr_ticker_count = 0
    
    Dim ticker_row As Long
    ticker_row = 2
    
    Dim ticker_row2
    ticker_row2 = ticker_row
    
    Dim current_ticker_name
    Dim max_ticker_name
    
    Dim x As Long
    
    Dim curr_open
    curr_open = Cells(2, 3).Value
    
    Dim curr_closed
    Dim curr_ticker_name As String
    
    Dim vol
    
    Dim opening As Long
    
    Dim closing As Long
    
    Dim new_opening
    
    Dim new_close
    
    Dim old_opening
    
    old_opening = ws.Cells(2, 3).Value
    
    Dim old_close
    
    Dim difference As Long
    
    Dim max_vol
    max_vol = 0
    
    Dim max_inc
    max_inc = 0
    
    Dim max_inc_ticker
    
    Dim max_dec
    max_dec = 0
    
    Dim max_dec_ticker
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    For row = 2 To LastRow
        If (ws.Cells(row, 1).Value = ws.Cells(row + 1, 1).Value) Then
            vol = vol + ws.Cells(row, 7).Value
        Else
            vol = vol + ws.Cells(row, 7).Value
            ws.Cells(ticker_row, 12).Value = vol
            
                       
            curr_ticker_name = ws.Cells(row, 1).Value
            ws.Cells(ticker_row, 9).Value = curr_ticker_name
            
            If (vol > max_vol) Then
                max_vol = vol
                max_ticker_name = curr_ticker_name
            End If
            
            
            ticker_row = ticker_row + 1
            vol = 0
        End If
    Next row
    
    ws.Cells(4, 17).Value = max_vol
    ws.Cells(4, 16).Value = max_ticker_name
    
    For x = 3 To (lastrow2)
                If (ws.Cells(x, 1).Value <> ws.Cells(x - 1, 1).Value) Then
                new_opening = ws.Cells(x, 3).Value
                old_close = ws.Cells(x - 1, 6)
                ws.Cells(ticker_row2, 10).Value = old_close - old_opening
                difference = old_close - old_opening
                        If (old_opening <> 0) Then
                        ws.Cells(ticker_row2, 11).Value = FormatPercent((difference / old_opening), 8)
                            If (ws.Cells(ticker_row2, 10).Value > 0) Then
                            ws.Cells(ticker_row2, 10).Interior.ColorIndex = 4
                            
                                If (ws.Cells(ticker_row2, 11).Value > max_inc) Then
                                    max_inc = ws.Cells(ticker_row2, 11).Value
                                    max_inc_ticker = ws.Cells(x, 1).Value
                                End If
                            
                            End If
                            
                            If (ws.Cells(ticker_row2, 10).Value < 0) Then
                            ws.Cells(ticker_row2, 10).Interior.ColorIndex = 3
                            
                                If (ws.Cells(ticker_row2, 11).Value < max_dec) Then
                                    max_dec = ws.Cells(ticker_row2, 11).Value
                                    max_dec_ticker = ws.Cells(x, 1).Value
                                End If
                            
                            
                            
                            End If
                        End If
                    ticker_row2 = ticker_row2 + 1
                    
                    old_opening = new_opening
                
                End If
                
    Next x
    
    If (Cells(ticker_row2, 10) < 0) Then
    ws.Cells(ticker_row2, 10).Interior.ColorIndex = 3
    End If
    
    If (Cells(ticker_row2, 10) > 0) Then
    ws.Cells(ticker_row2, 10).Interior.ColorIndex = 4
    End If
    
    ws.Cells(2, 16).Value = max_inc_ticker
    ws.Cells(2, 17).Value = FormatPercent(max_inc, 8)
    
    ws.Cells(3, 16).Value = max_dec_ticker
    ws.Cells(3, 17).Value = FormatPercent(max_dec, 8)
    
Next ws
End Sub
