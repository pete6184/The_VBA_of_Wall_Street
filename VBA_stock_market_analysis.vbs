Attribute VB_Name = "Module1"
Sub stock_market():

    Dim i, j, rowcount, rowcount2, greatest_volume As LongLong
    Dim ticker_symbol As String
    Dim yearly_change, opening_value, closing_value, percent_change, greatest_increase, greatest_decrease As Double
    Dim ws As Worksheet

    For Each ws In Sheets

    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    

    rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_symbol = 1
    yearly_change = 0
    percent_change = 0
    total_volume = 0
    opening_value = ws.Cells(2, 3).Value
    closing_value = 0

    For i = 2 To rowcount

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker_symbol = ticker_symbol + 1
            ws.Cells(ticker_symbol, 9).Value = ws.Cells(i, 1)

            total_volume = total_volume + ws.Cells(i, 7).Value
            ws.Cells(ticker_symbol, 12).Value = total_volume
            total_volume = 0

            closing_value = ws.Cells(i, 6).Value
            yearly_change = closing_value - opening_value
            ws.Cells(ticker_symbol, 10).Value = yearly_change
                If opening_value <> 0 Then
                    percent_change = yearly_change / opening_value
                    ws.Cells(ticker_symbol, 11).Value = percent_change
                    ws.Cells(ticker_symbol, 11).Value = FormatPercent(ws.Cells(ticker_symbol, 11))
                ElseIf closing_value = 0 Then
                    ws.Cells(ticker_symbol, 11).Value = "0"
                End If

                opening_value = ws.Cells(i + 1, 3).Value

        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            total_volume = total_volume + ws.Cells(i, 7).Value

        End If

    Next i
      
    rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    rowcount2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
    greatest_volume = WorksheetFunction.Max(ws.Range("L2:L" & rowcount).Value)
    greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & rowcount).Value)
    greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & rowcount).Value)
    
    For j = 2 To rowcount2
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 6
        End If
   
        If ws.Cells(j, 11).Value = greatest_increase Then
            ws.Cells(2, 16).Value = greatest_increase
            ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
            ws.Cells(2, 16).Value = FormatPercent(ws.Cells(2, 16))
        End If
        
        If ws.Cells(j, 11).Value = greatest_decrease Then
            ws.Cells(3, 16).Value = greatest_decrease
            ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
            ws.Cells(3, 16).Value = FormatPercent(ws.Cells(3, 16))
        End If
        
        If ws.Cells(j, 12).Value = greatest_volume Then
            ws.Cells(4, 16).Value = greatest_volume
            ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
        End If
        
    Next j

ws.Range("A:P").EntireColumn.AutoFit

Next ws


End Sub

